"""
microsoft_365_emad.flows.email — Microsoft 365 email via Graph API.

Uses MSAL public client + Graph REST API. No python-o365 dependency.
"""

import asyncio
import base64
import logging
from pathlib import Path

import httpx

from microsoft_365_emad import nadella_api_calls_total, nadella_api_errors_total
from microsoft_365_emad.o365_client import graph_get, graph_patch, graph_post

_log = logging.getLogger("microsoft_365_emad")


async def read_messages(
    folder: str = "inbox",
    limit: int = 20,
    unread_only: bool = False,
    since_hours: int | None = None,
) -> str:
    """Read messages from a mailbox folder."""

    def _sync():
        params = {
            "$top": str(limit),
            "$select": "subject,from,receivedDateTime,isRead,bodyPreview",
            "$orderby": "receivedDateTime DESC",
        }
        filters = []
        if unread_only:
            filters.append("isRead eq false")
        if since_hours:
            from datetime import datetime, timedelta, timezone

            since = (
                datetime.now(timezone.utc) - timedelta(hours=since_hours)
            ).isoformat()
            filters.append(f"receivedDateTime ge {since}")
        if filters:
            params["$filter"] = " and ".join(filters)

        result = graph_get(f"/me/mailFolders/{folder}/messages", params)
        if "error" in result:
            return result["error"]

        messages = result.get("value", [])
        if not messages:
            return "No messages found."

        lines = [f"Found {len(messages)} message(s):"]
        for m in messages:
            sender = m.get("from", {}).get("emailAddress", {}).get("address", "?")
            unread = "[UNREAD] " if not m.get("isRead") else ""
            subj = m.get("subject", "(no subject)")
            lines.append(f"- {unread}**{subj}** from {sender}")
        return "\n".join(lines)

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="email", operation="read").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="email", error_type=type(exc).__name__
        ).inc()
        return f"Error reading messages: {exc}"


async def send_message(
    to: str,
    subject: str,
    body: str,
    cc: str | None = None,
    attachment_paths: list[str] | None = None,
) -> str:
    """Send an email message via Graph API."""

    def _sync():
        message = {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [{"emailAddress": {"address": to}}],
        }
        if cc:
            message["ccRecipients"] = [{"emailAddress": {"address": cc}}]

        if attachment_paths:
            attachments = []
            for path_str in attachment_paths:
                path = Path(path_str)
                if path.exists():
                    content = base64.b64encode(path.read_bytes()).decode()
                    attachments.append(
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": path.name,
                            "contentBytes": content,
                        }
                    )
            if attachments:
                message["attachments"] = attachments

        result = graph_post("/me/sendMail", {"message": message})
        if result and "error" in result:
            return result["error"]
        return f"Email sent to {to} with subject '{subject}'."

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="email", operation="send").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="email", error_type=type(exc).__name__
        ).inc()
        return f"Error sending email: {exc}"


async def search_messages(query_text: str, limit: int = 10) -> str:
    """Search messages by text."""

    def _sync():
        params = {
            "$top": str(limit),
            "$search": f'"{query_text}"',
            "$select": "subject,from,receivedDateTime",
        }
        result = graph_get("/me/messages", params)
        if "error" in result:
            return result["error"]

        messages = result.get("value", [])
        if not messages:
            return f"No messages matching '{query_text}'."

        lines = [f"Search results for '{query_text}' ({len(messages)}):"]
        for m in messages:
            sender = m.get("from", {}).get("emailAddress", {}).get("address", "?")
            lines.append(f"- **{m.get('subject', '?')}** from {sender}")
        return "\n".join(lines)

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="email", operation="search").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="email", error_type=type(exc).__name__
        ).inc()
        return f"Error searching messages: {exc}"


async def mark_as_read(message_subject: str) -> str:
    """Mark messages matching a subject as read."""

    def _sync():
        # Find messages matching subject
        params = {
            "$filter": f"contains(subject, '{message_subject}')",
            "$select": "id,isRead",
            "$top": "10",
        }
        result = graph_get("/me/messages", params)
        if "error" in result:
            return result["error"]

        messages = result.get("value", [])
        count = 0
        for m in messages:
            if not m.get("isRead"):
                graph_patch(f"/me/messages/{m['id']}", {"isRead": True})
                count += 1
        return f"Marked {count} message(s) as read matching '{message_subject}'."

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="email", operation="mark_read").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="email", error_type=type(exc).__name__
        ).inc()
        return f"Error marking messages as read: {exc}"


async def list_folders() -> str:
    """List mailbox folders."""

    def _sync():
        result = graph_get("/me/mailFolders", {"$select": "displayName,totalItemCount"})
        if "error" in result:
            return result["error"]

        folders = result.get("value", [])
        if not folders:
            return "No folders found."
        lines = ["Mailbox folders:"]
        for f in folders:
            lines.append(f"- {f['displayName']} ({f.get('totalItemCount', 0)} items)")
        return "\n".join(lines)

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="email", operation="list_folders").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="email", error_type=type(exc).__name__
        ).inc()
        return f"Error listing folders: {exc}"
