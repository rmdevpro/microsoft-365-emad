"""
microsoft_365_emad.register — eMAD entry point.
"""

from langgraph.graph import StateGraph

EMAD_NAME = "microsoft-365-emad"
DESCRIPTION = (
    "Microsoft 365 Service Broker — email, calendar, OneDrive access "
    "with OAuth2 token management via MSAL."
)


def build_graph(params: dict | None = None) -> StateGraph:
    """Return the compiled Microsoft 365 Imperator StateGraph."""
    from microsoft_365_emad.flows.imperator import build_imperator_graph

    return build_imperator_graph(params)
