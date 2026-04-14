"""
Microsoft 365 Service Broker eMAD.

Prometheus counters registered at import time.
Safe for module re-import (hot-reload).
"""

from prometheus_client import Counter, REGISTRY


def _get_or_create_counter(name, description, labelnames):
    """Get existing counter or create new one. Safe for reimport."""
    try:
        return Counter(name, description, labelnames)
    except ValueError:
        for collector in REGISTRY._names_to_collectors.values():
            if hasattr(collector, '_name') and collector._name == name:
                return collector
        try:
            REGISTRY.unregister(REGISTRY._names_to_collectors.get(name))
        except (KeyError, AttributeError):
            pass
        return Counter(name, description, labelnames)


m365_api_calls_total = _get_or_create_counter(
    "m365_api_calls_total",
    "Microsoft Graph API calls",
    ["service", "operation"],
)
m365_api_errors_total = _get_or_create_counter(
    "m365_api_errors_total",
    "Microsoft Graph API errors",
    ["service", "error_type"],
)
m365_token_refreshes_total = _get_or_create_counter(
    "m365_token_refreshes_total",
    "Token refresh operations",
    ["account"],
)
m365_token_refresh_failures_total = _get_or_create_counter(
    "m365_token_refresh_failures_total",
    "Failed token refreshes",
    ["account"],
)

from microsoft_365_emad.register import build_graph  # noqa: E402, F401
