"""Shared fixtures and helpers for the SecOps Portal test suite."""

from __future__ import annotations

import io
import json
import os

import pandas as pd
import pytest


# ─── Markers ──────────────────────────────────────────────────────────────────
# Tests marked @pytest.mark.integration are skipped unless N8N_WEBHOOK_URL
# is set in the environment. This lets the unit suite run offline / in CI.
def pytest_configure(config):
    config.addinivalue_line("markers", "unit: pure unit tests — no external deps")
    config.addinivalue_line(
        "markers",
        "integration: requires n8n running (set N8N_WEBHOOK_URL env var)",
    )


def pytest_collection_modifyitems(config, items):
    if not os.getenv("N8N_WEBHOOK_URL"):
        skip = pytest.mark.skip(reason="N8N_WEBHOOK_URL not set — skipping integration tests")
        for item in items:
            if "integration" in item.keywords:
                item.add_marker(skip)


# ─── Fixtures ─────────────────────────────────────────────────────────────────
@pytest.fixture
def n8n_webhook_url() -> str:
    """Return the live n8n webhook URL (only available in integration runs)."""
    url = os.environ.get("N8N_WEBHOOK_URL", "")
    if not url:
        pytest.skip("N8N_WEBHOOK_URL not set")
    return url


@pytest.fixture
def valid_vt_response() -> dict:
    """Minimal VirusTotal-shaped n8n response for a benign IP."""
    return {
        "status": "success",
        "target": "8.8.8.8",
        "threat_score": 2,
        "location": "United States",
        "known_malicious": False,
        "summary": "8.8.8.8 is Google's public DNS server. No threats detected.",
        "details": "2 of 91 VirusTotal engines flagged this target. ISP: Google LLC",
    }


@pytest.fixture
def malicious_response() -> dict:
    """Minimal n8n response for a known-bad IP."""
    return {
        "status": "success",
        "target": "198.51.100.1",
        "threat_score": 87,
        "location": "Russia",
        "known_malicious": True,
        "summary": "This IP is a known TOR exit node. Recommend blocking at perimeter.",
        "details": "79 of 91 VirusTotal engines flagged this target. ISP: AS666",
    }


@pytest.fixture
def csv_file_obj() -> io.BytesIO:
    """In-memory CSV file object that mimics a Streamlit UploadedFile."""
    content = b"Source IP,Destination IP,Port,Action\n192.168.1.1,8.8.8.8,53,Allowed\n"
    buf = io.BytesIO(content)
    buf.name = "logs.csv"
    return buf


@pytest.fixture
def xlsx_file_obj() -> io.BytesIO:
    """In-memory XLSX file object that mimics a Streamlit UploadedFile."""
    df = pd.DataFrame({"Source IP": ["10.0.0.1"], "Port": [443], "Action": ["Blocked"]})
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    buf.name = "logs.xlsx"
    return buf


@pytest.fixture
def bad_file_obj() -> io.BytesIO:
    """Garbage bytes with an unsupported extension — should return None."""
    buf = io.BytesIO(b"\x00\x01\x02garbage")
    buf.name = "data.parquet"
    return buf
