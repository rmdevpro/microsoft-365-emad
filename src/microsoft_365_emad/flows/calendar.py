"""
microsoft_365_emad.flows.calendar — Microsoft 365 calendar via Graph API.
"""

import asyncio
import logging

import httpx

from microsoft_365_emad import nadella_api_calls_total, nadella_api_errors_total
from microsoft_365_emad.o365_client import graph_delete, graph_get, graph_post

_log = logging.getLogger("microsoft_365_emad")


async def list_events(days_ahead: int = 7, limit: int = 20) -> str:
    """List upcoming calendar events."""

    def _sync():
        from datetime import datetime, timedelta, timezone

        now = datetime.now(timezone.utc)
        end = now + timedelta(days=days_ahead)

        params = {
            "$top": str(limit),
            "$select": "subject,start,end,location,organizer",
            "$orderby": "start/dateTime",
            "$filter": f"start/dateTime ge '{now.isoformat()}' and start/dateTime le '{end.isoformat()}'",
        }
        result = graph_get("/me/events", params)
        if "error" in result:
            return result["error"]

        events = result.get("value", [])
        if not events:
            return f"No events in the next {days_ahead} days."

        lines = [f"Events for the next {days_ahead} days:"]
        for e in events:
            start = e.get("start", {}).get("dateTime", "?")[:16]
            subj = e.get("subject", "(no title)")
            loc = e.get("location", {}).get("displayName", "")
            loc_str = f" at {loc}" if loc else ""
            lines.append(f"- **{subj}** {start}{loc_str}")
        return "\n".join(lines)

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="calendar", operation="list").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="calendar", error_type=type(exc).__name__
        ).inc()
        return f"Error listing events: {exc}"


async def create_event(
    subject: str,
    start: str,
    end: str,
    body: str = "",
    location: str = "",
    attendees: list[str] | None = None,
) -> str:
    """Create a calendar event."""

    def _sync():
        event = {
            "subject": subject,
            "start": {"dateTime": start, "timeZone": "UTC"},
            "end": {"dateTime": end, "timeZone": "UTC"},
        }
        if body:
            event["body"] = {"contentType": "Text", "content": body}
        if location:
            event["location"] = {"displayName": location}
        if attendees:
            event["attendees"] = [
                {"emailAddress": {"address": a}, "type": "required"} for a in attendees
            ]

        result = graph_post("/me/events", event)
        if result and "error" in result:
            return result["error"]
        event_id = result.get("id", "?") if result else "?"
        return f"Created event '{subject}' (id: {event_id})."

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="calendar", operation="create").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="calendar", error_type=type(exc).__name__
        ).inc()
        return f"Error creating event: {exc}"


async def delete_event(event_subject: str) -> str:
    """Delete a calendar event by subject match."""

    def _sync():
        params = {
            "$filter": f"contains(subject, '{event_subject}')",
            "$select": "id,subject",
            "$top": "10",
        }
        result = graph_get("/me/events", params)
        if "error" in result:
            return result["error"]

        events = result.get("value", [])
        for e in events:
            if event_subject.lower() in e.get("subject", "").lower():
                graph_delete(f"/me/events/{e['id']}")
                return f"Deleted event '{e['subject']}'."
        return f"No event found matching '{event_subject}'."

    try:
        result = await asyncio.to_thread(_sync)
        nadella_api_calls_total.labels(service="calendar", operation="delete").inc()
        return result
    except (httpx.HTTPError, RuntimeError, OSError) as exc:
        nadella_api_errors_total.labels(
            service="calendar", error_type=type(exc).__name__
        ).inc()
        return f"Error deleting event: {exc}"
