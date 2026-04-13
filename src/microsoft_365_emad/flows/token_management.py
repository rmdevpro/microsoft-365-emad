"""
microsoft_365_emad.flows.token_management — OAuth2 token health and refresh.
"""

import logging

from microsoft_365_emad import (
    m365_token_refresh_failures_total,
    m365_token_refreshes_total,
)
from microsoft_365_emad.o365_client import check_token_health, ensure_authenticated

_log = logging.getLogger("microsoft_365_emad")


async def refresh_all_tokens() -> str:
    """Check and refresh tokens for all configured accounts.

    Called by the scheduler every 5 minutes.
    python-o365 handles the actual refresh internally when
    account.is_authenticated checks the token.
    """
    # For now, check the known accounts
    # In the future, this could read from the tokens table
    import os

    accounts_str = os.environ.get("M365_ACCOUNTS", "")
    if not accounts_str:
        return "No accounts configured (set M365_ACCOUNTS env var)."

    accounts = [a.strip() for a in accounts_str.split(",") if a.strip()]
    results = []

    for username in accounts:
        health = await check_token_health(username)
        if health["authenticated"]:
            m365_token_refreshes_total.labels(account=username).inc()
            results.append(f"{username}: healthy")
        else:
            # Try to re-authenticate (will use stored refresh token)
            success, msg = await ensure_authenticated(username)
            if success:
                m365_token_refreshes_total.labels(account=username).inc()
                results.append(f"{username}: refreshed")
            else:
                m365_token_refresh_failures_total.labels(account=username).inc()
                results.append(f"{username}: NEEDS RE-CONSENT — {msg}")

    return "\n".join(results)


async def get_token_status() -> str:
    """Report on the health of all managed tokens."""
    import os

    accounts_str = os.environ.get("M365_ACCOUNTS", "")
    if not accounts_str:
        return "No accounts configured."

    accounts = [a.strip() for a in accounts_str.split(",") if a.strip()]
    lines = ["Token status:"]

    for username in accounts:
        health = await check_token_health(username)
        status = "authenticated" if health["authenticated"] else "NOT AUTHENTICATED"
        lines.append(f"  {username}: {status}")

    return "\n".join(lines)
