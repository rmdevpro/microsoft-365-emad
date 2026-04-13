"""
Microsoft 365 Service Broker eMAD.

Prometheus counters registered at import time.
"""

from prometheus_client import Counter

m365_api_calls_total = Counter(
    "m365_api_calls_total",
    "Microsoft Graph API calls",
    ["service", "operation"],
)
m365_api_errors_total = Counter(
    "m365_api_errors_total",
    "Microsoft Graph API errors",
    ["service", "error_type"],
)
m365_token_refreshes_total = Counter(
    "m365_token_refreshes_total",
    "Token refresh operations",
    ["account"],
)
m365_token_refresh_failures_total = Counter(
    "m365_token_refresh_failures_total",
    "Failed token refreshes",
    ["account"],
)

from microsoft_365_emad.register import build_graph  # noqa: E402, F401
