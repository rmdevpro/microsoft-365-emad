"""
microsoft_365_emad.o365_client — Microsoft Graph API client via MSAL.

Uses the public client ID trick — Microsoft's own Graph CLI client ID
(14d82eec-204b-4c2f-b7e8-296a70dab67e) with device code flow. No app
registration needed. No client secret. Works with personal O365 accounts.

Token cached via MSAL's SerializableTokenCache to a JSON file.
Auto-refreshes silently on each call. Device code flow only needed once
(or when refresh token expires after 90 days).
"""

import json
import logging
import os
from pathlib import Path

import httpx
import msal

_log = logging.getLogger("microsoft_365_emad")

# Microsoft's own Graph CLI public client ID — no app registration needed
_PUBLIC_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
_AUTHORITY = "https://login.microsoftonline.com/consumers"
_GRAPH_BASE = "https://graph.microsoft.com/v1.0"

_TOKEN_CACHE_PATH = Path(
    os.environ.get(
        "M365_TOKEN_CACHE",
        "/storage/credentials/microsoft-365/msal_token_cache.json",
    )
)

_DEFAULT_SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Files.ReadWrite.All",
]

# Module-level cache for the MSAL app + token cache
_msal_app: msal.PublicClientApplication | None = None
_token_cache: msal.SerializableTokenCache | None = None


def _get_msal_app() -> msal.PublicClientApplication:
    """Get or create the MSAL public client app with persistent token cache."""
    global _msal_app, _token_cache

    if _msal_app is not None:
        return _msal_app

    _token_cache = msal.SerializableTokenCache()

    # Load existing cache from file
    if _TOKEN_CACHE_PATH.exists():
        try:
            _token_cache.deserialize(_TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
            _log.info("Loaded MSAL token cache from %s", _TOKEN_CACHE_PATH)
        except (json.JSONDecodeError, OSError) as exc:
            _log.warning("Failed to load token cache: %s", exc)

    _msal_app = msal.PublicClientApplication(
        _PUBLIC_CLIENT_ID,
        authority=_AUTHORITY,
        token_cache=_token_cache,
    )
    return _msal_app


def _save_cache() -> None:
    """Persist the MSAL token cache to disk."""
    if _token_cache is not None and _token_cache.has_state_changed:
        _TOKEN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        _TOKEN_CACHE_PATH.write_text(_token_cache.serialize(), encoding="utf-8")


def get_access_token(scopes: list[str] | None = None) -> str | None:
    """Get a valid access token, silently refreshing if needed.

    Returns the access token string, or None if not authenticated.
    """
    if scopes is None:
        scopes = _DEFAULT_SCOPES

    app = _get_msal_app()

    # Try silent token acquisition first (uses cached refresh token)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result and "access_token" in result:
            _save_cache()
            return result["access_token"]

    return None


def initiate_device_code_flow(
    scopes: list[str] | None = None,
) -> dict:
    """Start the device code flow. Returns the flow dict with 'message' key.

    The caller should display flow['message'] to the user, which contains
    the URL and code they need to enter.
    """
    if scopes is None:
        scopes = _DEFAULT_SCOPES

    app = _get_msal_app()
    flow = app.initiate_device_flow(scopes=scopes)
    if "message" not in flow:
        raise RuntimeError(f"Device code flow failed: {flow}")
    return flow


def complete_device_code_flow(flow: dict) -> tuple[bool, str]:
    """Complete the device code flow (blocking — waits for user to authenticate).

    Returns (success, message).
    """
    app = _get_msal_app()
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        _save_cache()
        _log.info("Device code authentication successful")
        return True, "Authentication successful."

    error = result.get("error_description", result.get("error", "Unknown error"))
    return False, f"Authentication failed: {error}"


def is_authenticated() -> bool:
    """Check if we have a valid (or refreshable) token."""
    return get_access_token() is not None


def graph_get(endpoint: str, params: dict | None = None) -> dict:
    """Make a GET request to Microsoft Graph API.

    Args:
        endpoint: Graph API path (e.g., '/me/messages').
        params: Query parameters.

    Returns:
        JSON response as dict.
    """
    token = get_access_token()
    if not token:
        return {"error": "Not authenticated. Run device code flow first."}

    url = f"{_GRAPH_BASE}{endpoint}"
    with httpx.Client(timeout=30.0) as client:
        resp = client.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params=params,
        )
        if resp.status_code == 401:
            return {"error": "Token expired or revoked. Re-authenticate."}
        resp.raise_for_status()
        return resp.json()


def graph_post(endpoint: str, body: dict) -> dict | None:
    """Make a POST request to Microsoft Graph API.

    Args:
        endpoint: Graph API path (e.g., '/me/sendMail').
        body: JSON body.

    Returns:
        JSON response as dict, or None for 202/204 responses.
    """
    token = get_access_token()
    if not token:
        return {"error": "Not authenticated. Run device code flow first."}

    url = f"{_GRAPH_BASE}{endpoint}"
    with httpx.Client(timeout=30.0) as client:
        resp = client.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=body,
        )
        if resp.status_code == 401:
            return {"error": "Token expired or revoked. Re-authenticate."}
        resp.raise_for_status()
        if resp.status_code in (202, 204):
            return None
        return resp.json()


def graph_patch(endpoint: str, body: dict) -> dict | None:
    """Make a PATCH request to Microsoft Graph API."""
    token = get_access_token()
    if not token:
        return {"error": "Not authenticated. Run device code flow first."}

    url = f"{_GRAPH_BASE}{endpoint}"
    with httpx.Client(timeout=30.0) as client:
        resp = client.patch(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=body,
        )
        if resp.status_code == 401:
            return {"error": "Token expired or revoked. Re-authenticate."}
        resp.raise_for_status()
        if resp.status_code in (202, 204):
            return None
        return resp.json()


def graph_delete(endpoint: str) -> bool:
    """Make a DELETE request to Microsoft Graph API. Returns True on success."""
    token = get_access_token()
    if not token:
        return False

    url = f"{_GRAPH_BASE}{endpoint}"
    with httpx.Client(timeout=30.0) as client:
        resp = client.delete(
            url,
            headers={"Authorization": f"Bearer {token}"},
        )
        resp.raise_for_status()
        return True
