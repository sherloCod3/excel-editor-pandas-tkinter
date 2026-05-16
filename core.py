import random
import time
from typing import Any
from urllib.parse import urlparse

import pandas as pd
import requests


# Required keys in a valid n8n SOC response.
_REQUIRED_KEYS: set[str] = {"status", "target", "threat_score", "location", "known_malicious"}

# Hostname prefixes that must never be the target of an outbound webhook request.
# Allowing them would let a user pivot the server into probing internal infrastructure.
_SSRF_BLOCKED_HOSTS = ("localhost", "127.", "0.0.0.0", "169.254.", "::1", "[::1]")


def _validate_webhook_url(url: str) -> bool:
    """Return True only if the URL is a safe external endpoint.

    Blocks loopback addresses, link-local ranges (AWS/GCP/Azure instance
    metadata at 169.254.169.254), and non-HTTP(S) schemes to prevent SSRF.
    """
    try:
        parsed = urlparse(url)
    except ValueError:
        return False
    if parsed.scheme not in {"http", "https"}:
        return False
    host = parsed.hostname or ""
    return not any(host.startswith(blocked) for blocked in _SSRF_BLOCKED_HOSTS)


def _validate_response(data: dict[str, Any]) -> dict[str, Any]:
    """Ensure the n8n response matches the expected schema.

    Fills in safe defaults for any missing fields and casts `threat_score`
    to int so the UI never receives unexpected types.
    """
    defaults: dict[str, Any] = {
        "status": "success",
        "target": "",
        "threat_score": 0,
        "location": "Unknown",
        "known_malicious": False,
        "summary": "",
        "details": "",
    }
    # Merge: n8n data wins over defaults.
    merged = {**defaults, **data}
    # Guarantee threat_score is always an int clamped to [0, 100].
    try:
        merged["threat_score"] = max(0, min(100, int(merged["threat_score"])))
    except (ValueError, TypeError):
        merged["threat_score"] = 0
    return merged


def load_data(file) -> pd.DataFrame | None:
    """Parse an uploaded CSV or Excel file into a DataFrame."""
    try:
        if file.name.endswith(".csv"):
            return pd.read_csv(file)
        elif file.name.endswith((".xls", ".xlsx")):
            return pd.read_excel(file)
        return None
    except Exception:
        return None


def fetch_soc_data(target_url: str, webhook_url: str | None = None) -> dict[str, Any]:
    """Call the n8n SOC webhook or return mock data when no URL is configured.

    Args:
        target_url:  The IP address or URL/domain to investigate.
        webhook_url: The n8n webhook endpoint. When empty/None the function
                     runs in mock mode — useful for development and demos.

    Returns:
        A dict matching the SOC response schema:
        { status, target, threat_score, location, known_malicious, summary, details }
    """
    # Mock mode: return randomized data without hitting n8n.
    if not webhook_url:
        time.sleep(1.5)
        return _validate_response({
            "status": "success",
            "target": target_url,
            "threat_score": random.randint(10, 95),
            "location": random.choice(["United States", "Russia", "China", "Brazil", "Germany"]),
            "known_malicious": random.choice([True, False]),
            "summary": "",
            "details": "Simulated scan — configure the webhook URL for real threat data.",
        })

    # Live mode: POST to the n8n webhook and validate the response.
    if not _validate_webhook_url(webhook_url):
        return {"error": f"Blocked request to disallowed host: {webhook_url!r}"}

    try:
        response = requests.post(
            webhook_url,
            json={"target": target_url},
            timeout=15,  # VirusTotal P99 can be slow
        )
        response.raise_for_status()
        raw: dict[str, Any] = response.json()
        return _validate_response(raw)
    except requests.exceptions.RequestException as exc:
        return {"error": str(exc)}
