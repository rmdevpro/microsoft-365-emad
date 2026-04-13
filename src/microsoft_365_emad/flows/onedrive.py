"""
microsoft_365_emad.flows.onedrive — Microsoft 365 OneDrive via Graph API.
"""

import asyncio
import logging
from pathlib import Path

import httpx

from microsoft_365_emad import nadella_api_calls_total, nadella_api_errors_total
from microsoft_365_emad.o365_client import graph_delete, graph_get, graph_post

_log = logging.getLogger("microsoft_365_emad")


async def list_files(path: str = "/", limit: int = 50) -> str:
    """List files and folders at a OneDrive path."""

    def _sync():
        if path == "/":
            endpoint = "/me/drive/root/children"
        else:
            endpoint = f"/me/drive/root:/{path.strip('/')}:/children"

        result = graph_get(
            endpoint,
            {
                "$top": str(limit),
                "$select": "name,size,folder,file,lastModifiedDateTime",
            },
        )
        if "error" in result:
            return result["error"]

        items = result.get("value", [])
        if not items:
            return f"No items at {path}."

        lines = [f"OneDrive {path}:"]
        for item in items:
            icon = "[dir]" if "folder" in item else "[file]"
            size = f" ({item.get('size', 0)} bytes)" if "file" in item else ""
            lines.append(f"  {icon} {item['name']}{size}")
        return "\n".join(lines)

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="onedrive", operation="list").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="onedrive", error_type=type(exc).__name__
        ).inc()
        return f"Error listing files: {exc}"


async def upload_file(local_path: str, remote_path: str) -> str:
    """Upload a file to OneDrive."""

    def _sync():
        local = Path(local_path)
        if not local.exists():
            return f"Local file not found: {local_path}"

        # For small files (< 4MB), use simple upload
        content = local.read_bytes()
        if len(content) > 4 * 1024 * 1024:
            return "File too large for simple upload (> 4MB). Use upload session."

        from microsoft_365_emad.o365_client import get_access_token

        token = get_access_token()
        if not token:
            return "Not authenticated."

        remote = remote_path.strip("/")
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{remote}:/content"

        with httpx.Client(timeout=60.0) as client:
            resp = client.put(
                url,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/octet-stream",
                },
                content=content,
            )
            resp.raise_for_status()

        return f"Uploaded {local.name} to OneDrive /{remote}."

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="onedrive", operation="upload").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="onedrive", error_type=type(exc).__name__
        ).inc()
        return f"Error uploading file: {exc}"


async def search_files(query: str, limit: int = 20) -> str:
    """Search OneDrive files."""

    def _sync():
        result = graph_get(
            f"/me/drive/root/search(q='{query}')",
            {"$top": str(limit), "$select": "name,size,webUrl"},
        )
        if "error" in result:
            return result["error"]

        items = result.get("value", [])
        if not items:
            return f"No files matching '{query}'."

        lines = [f"Search results for '{query}':"]
        for item in items:
            lines.append(f"  {item['name']}")
        return "\n".join(lines)

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="onedrive", operation="search").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="onedrive", error_type=type(exc).__name__
        ).inc()
        return f"Error searching files: {exc}"


async def create_folder(name: str, path: str = "/") -> str:
    """Create a folder on OneDrive."""

    def _sync():
        if path == "/":
            endpoint = "/me/drive/root/children"
        else:
            endpoint = f"/me/drive/root:/{path.strip('/')}:/children"

        body = {
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename",
        }
        result = graph_post(endpoint, body)
        if result and "error" in result:
            return result["error"]
        return f"Created folder '{name}' at {path}."

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(
            service="onedrive", operation="create_folder"
        ).inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="onedrive", error_type=type(exc).__name__
        ).inc()
        return f"Error creating folder: {exc}"


async def delete_item(item_path: str) -> str:
    """Delete a file or folder from OneDrive."""

    def _sync():
        remote = item_path.strip("/")
        success = graph_delete(f"/me/drive/root:/{remote}")
        if success:
            return f"Deleted '/{remote}' from OneDrive."
        return f"Failed to delete '/{remote}'."

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="onedrive", operation="delete").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="onedrive", error_type=type(exc).__name__
        ).inc()
        return f"Error deleting item: {exc}"
