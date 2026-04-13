"""
Nadella — Microsoft 365 Service Broker eMAD.

Prometheus counters registered at import time.
"""

from prometheus_client import Counter

nadella_api_calls_total = Counter(
    "nadella_api_calls_total",
    "Microsoft Graph API calls",
    ["service", "operation"],
)
nadella_api_errors_total = Counter(
    "nadella_api_errors_total",
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
