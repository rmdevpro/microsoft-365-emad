"""
microsoft_365_emad.flows.imperator — Nadella Imperator StateGraph.

CB-style ReAct agent for Microsoft 365 service brokering.
Uses MSAL public client + Graph API. No app registration needed.
"""

import logging
from pathlib import Path
from typing import Annotated, TypedDict

import httpx
import openai
from langchain_core.messages import AIMessage, AnyMessage, SystemMessage, ToolMessage
from langchain_core.tools import tool
from langgraph.graph import END, StateGraph
from langgraph.graph.message import add_messages
from langgraph.prebuilt import ToolNode

_log = logging.getLogger("microsoft_365_emad")

_PROMPT_PATH = (
    Path(__file__).resolve().parent.parent / "prompts" / "imperator_identity.md"
)
_MAX_ITERATIONS = 15


class NadellaImperatorState(TypedDict):
    payload: dict  # Full OpenAI request body
    messages: Annotated[list[AnyMessage], add_messages]
    conversation_id: str | None
    response_text: str | None
    error: str | None
    iteration_count: int


# ── Tools ────────────────────────────────────────────────────────────────


@tool
async def read_email(
    folder: str = "inbox",
    limit: int = 20,
    unread_only: bool = False,
    since_hours: int = 0,
) -> str:
    """Read email messages from the O365 mailbox.

    Args:
        folder: Mailbox folder (default: inbox).
        limit: Max messages to return.
        unread_only: Only return unread messages.
        since_hours: Only return messages from the last N hours (0 = no filter).
    """
    from microsoft_365_emad.flows.email import read_messages

    return await read_messages(folder, limit, unread_only, since_hours or None)


@tool
async def send_email(
    to: str,
    subject: str,
    body: str,
    cc: str = "",
    attachment_paths: str = "",
) -> str:
    """Send an email from the O365 account.

    Args:
        to: Recipient email address.
        subject: Email subject.
        body: Email body text.
        cc: CC recipient (optional).
        attachment_paths: Comma-separated file paths to attach (optional).
    """
    from microsoft_365_emad.flows.email import send_message

    attachments = (
        [p.strip() for p in attachment_paths.split(",") if p.strip()]
        if attachment_paths
        else None
    )
    return await send_message(to, subject, body, cc or None, attachments)


@tool
async def search_email(query: str, limit: int = 10) -> str:
    """Search email messages by text.

    Args:
        query: Search text.
        limit: Max results.
    """
    from microsoft_365_emad.flows.email import search_messages

    return await search_messages(query, limit)


@tool
async def list_email_folders() -> str:
    """List mailbox folders."""
    from microsoft_365_emad.flows.email import list_folders

    return await list_folders()


@tool
async def mark_email_read(message_subject: str) -> str:
    """Mark emails as read by subject match.

    Args:
        message_subject: Subject line to match.
    """
    from microsoft_365_emad.flows.email import mark_as_read

    return await mark_as_read(message_subject)


@tool
async def list_calendar_events(days_ahead: int = 7, limit: int = 20) -> str:
    """List upcoming calendar events.

    Args:
        days_ahead: How many days ahead to look.
        limit: Max events.
    """
    from microsoft_365_emad.flows.calendar import list_events

    return await list_events(days_ahead, limit)


@tool
async def create_calendar_event(
    subject: str,
    start: str,
    end: str,
    body: str = "",
    location: str = "",
    attendees: str = "",
) -> str:
    """Create a calendar event.

    Args:
        subject: Event subject.
        start: Start time (ISO format).
        end: End time (ISO format).
        body: Event body (optional).
        location: Event location (optional).
        attendees: Comma-separated attendee emails (optional).
    """
    from microsoft_365_emad.flows.calendar import create_event

    attendee_list = (
        [a.strip() for a in attendees.split(",") if a.strip()] if attendees else None
    )
    return await create_event(subject, start, end, body, attendee_list, location)


@tool
async def delete_calendar_event(event_subject: str) -> str:
    """Delete a calendar event by subject.

    Args:
        event_subject: Event subject to match.
    """
    from microsoft_365_emad.flows.calendar import delete_event

    return await delete_event(event_subject)


@tool
async def list_onedrive_files(path: str = "/", limit: int = 50) -> str:
    """List files and folders on OneDrive.

    Args:
        path: OneDrive path (default: root).
        limit: Max items.
    """
    from microsoft_365_emad.flows.onedrive import list_files

    return await list_files(path, limit)


@tool
async def upload_to_onedrive(local_path: str, remote_path: str) -> str:
    """Upload a file to OneDrive.

    Args:
        local_path: Local file path.
        remote_path: OneDrive destination path.
    """
    from microsoft_365_emad.flows.onedrive import upload_file

    return await upload_file(local_path, remote_path)


@tool
async def search_onedrive(query: str, limit: int = 20) -> str:
    """Search OneDrive files.

    Args:
        query: Search text.
        limit: Max results.
    """
    from microsoft_365_emad.flows.onedrive import search_files

    return await search_files(query, limit)


@tool
async def create_onedrive_folder(name: str, path: str = "/") -> str:
    """Create a folder on OneDrive.

    Args:
        name: Folder name.
        path: Parent path (default: root).
    """
    from microsoft_365_emad.flows.onedrive import create_folder

    return await create_folder(name, path)


@tool
async def delete_onedrive_item(item_path: str) -> str:
    """Delete a file or folder from OneDrive.

    Args:
        item_path: Path to the item.
    """
    from microsoft_365_emad.flows.onedrive import delete_item

    return await delete_item(item_path)


@tool
async def authenticate() -> str:
    """Start authentication for the O365 account. Returns device code instructions."""
    import asyncio

    from microsoft_365_emad.o365_client import (
        complete_device_code_flow,
        initiate_device_code_flow,
        is_authenticated,
    )

    if is_authenticated():
        return "Already authenticated."

    flow = initiate_device_code_flow()
    message = flow["message"]

    # The device code flow blocks until the user authenticates
    # In headless Kaiser, we return the message and let the user complete it
    # Then complete the flow in a background thread
    success, result_msg = await asyncio.to_thread(complete_device_code_flow, flow)
    if success:
        return (
            f"Authentication successful.\n\nDevice code instructions were:\n{message}"
        )
    return f"Authentication required.\n\n{message}\n\n(Complete the steps above, then try again.)"


@tool
async def check_token_status() -> str:
    """Check if the O365 account is authenticated."""
    from microsoft_365_emad.o365_client import is_authenticated

    if is_authenticated():
        return "Authenticated — token is valid."
    return "Not authenticated — device code flow needed."


_TOOLS = [
    read_email,
    send_email,
    search_email,
    list_email_folders,
    mark_email_read,
    list_calendar_events,
    create_calendar_event,
    delete_calendar_event,
    list_onedrive_files,
    upload_to_onedrive,
    search_onedrive,
    create_onedrive_folder,
    delete_onedrive_item,
    authenticate,
    check_token_status,
]


# ── System prompt ────────────────────────────────────────────────────────


def _load_system_prompt() -> str:
    if _PROMPT_PATH.exists():
        return _PROMPT_PATH.read_text(encoding="utf-8")
    return "You are Nadella, the Microsoft 365 service broker."


# ── Graph nodes ──────────────────────────────────────────────────────────


async def init_node(state: NadellaImperatorState) -> dict:
    """Parse payload, set up conversation. Append-only on resumed turns."""
    import uuid
    from langchain_core.messages import HumanMessage

    payload = state.get("payload", {})
    existing_messages = state.get("messages", [])

    conv_id = payload.get("conversation_id")
    if conv_id == "new":
        conv_id = str(uuid.uuid4())
    elif not conv_id:
        conv_id = f"default-{payload.get('model', 'nadella')}"

    # Extract last user message
    raw_messages = payload.get("messages", [])
    new_user_msg = None
    for m in reversed(raw_messages):
        if m.get("role") == "user":
            new_user_msg = HumanMessage(content=m.get("content", ""))
            break
    if not new_user_msg:
        new_user_msg = HumanMessage(content="")

    # Resumed conversation
    if existing_messages:
        return {"messages": [new_user_msg], "conversation_id": conv_id, "iteration_count": 0}

    # First turn
    system_content = _load_system_prompt()
    messages = [SystemMessage(content=system_content), new_user_msg]
    return {"messages": messages, "conversation_id": conv_id, "iteration_count": 0}


async def agent_node(state: NadellaImperatorState) -> dict:
    from microsoft_365_emad.inference import get_llm

    llm = get_llm("fast")
    llm_with_tools = llm.bind_tools(_TOOLS)

    messages = list(state["messages"])

    max_messages = 30
    if len(messages) > max_messages:
        cut_index = len(messages) - (max_messages - 1)
        while cut_index < len(messages) and isinstance(
            messages[cut_index], ToolMessage
        ):
            cut_index += 1
        messages = [messages[0]] + messages[cut_index:]

    try:
        response = await llm_with_tools.ainvoke(messages)
    except (openai.APIError, httpx.HTTPError, ValueError, RuntimeError) as exc:
        _log.error("Nadella LLM call failed: %s", exc, exc_info=True)
        return {
            "messages": [
                AIMessage(content="I encountered an error processing your request.")
            ],
            "error": str(exc),
        }

    return {
        "messages": [response],
        "iteration_count": state.get("iteration_count", 0) + 1,
    }


def should_continue(state: NadellaImperatorState) -> str:
    if state.get("error"):
        return "finalize"
    messages = state["messages"]
    if not messages:
        return "finalize"
    last = messages[-1]
    if isinstance(last, AIMessage) and last.tool_calls:
        if state.get("iteration_count", 0) >= _MAX_ITERATIONS:
            return "finalize"
        return "tool_node"
    return "finalize"


def finalize(state: NadellaImperatorState) -> dict:
    for msg in reversed(state["messages"]):
        if (
            isinstance(msg, AIMessage)
            and msg.content
            and not getattr(msg, "tool_calls", None)
        ):
            return {
                "response_text": str(msg.content),
                "conversation_id": state.get("conversation_id"),
            }
    return {
        "response_text": "[No response generated]",
        "conversation_id": state.get("conversation_id"),
    }


# ── Graph builder ────────────────────────────────────────────────────────


def build_imperator_graph(params: dict | None = None) -> StateGraph:
    tool_node_instance = ToolNode(_TOOLS)

    workflow = StateGraph(NadellaImperatorState)
    workflow.add_node("init_node", init_node)
    workflow.add_node("agent_node", agent_node)
    workflow.add_node("tool_node", tool_node_instance)
    workflow.add_node("finalize", finalize)

    workflow.set_entry_point("init_node")
    workflow.add_edge("init_node", "agent_node")
    workflow.add_conditional_edges(
        "agent_node",
        should_continue,
        {"tool_node": "tool_node", "finalize": "finalize"},
    )
    workflow.add_edge("tool_node", "agent_node")
    workflow.add_edge("finalize", END)

    from app.checkpointer import get_checkpointer

    return workflow.compile(checkpointer=get_checkpointer())
