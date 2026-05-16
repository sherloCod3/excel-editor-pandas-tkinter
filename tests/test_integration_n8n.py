"""Integration tests — require a live n8n instance.

Run with:
    N8N_WEBHOOK_URL=http://localhost:5678/webhook/soc-scan pytest -m integration

These tests are automatically skipped when N8N_WEBHOOK_URL is not set,
so the unit suite always passes in CI without Docker.
"""

import pytest
import requests
from core import _validate_response


@pytest.mark.integration
class TestN8nWebhookContract:
    """Verify the live n8n workflow returns a response matching the schema
    expected by the Streamlit frontend."""

    def _post(self, url: str, target: str) -> dict:
        resp = requests.post(url, json={"target": target}, timeout=20)
        resp.raise_for_status()
        return resp.json()

    def test_webhook_is_reachable(self, n8n_webhook_url):
        resp = requests.post(n8n_webhook_url, json={"target": "8.8.8.8"}, timeout=20)
        assert resp.status_code == 200

    def test_response_is_json(self, n8n_webhook_url):
        resp = requests.post(n8n_webhook_url, json={"target": "8.8.8.8"}, timeout=20)
        assert isinstance(resp.json(), dict)

    def test_required_keys_present(self, n8n_webhook_url):
        data = self._post(n8n_webhook_url, "8.8.8.8")
        required = {"status", "target", "threat_score", "location", "known_malicious"}
        assert not required - data.keys(), f"Missing keys: {required - data.keys()}"

    def test_threat_score_is_int_in_range(self, n8n_webhook_url):
        data = self._post(n8n_webhook_url, "8.8.8.8")
        score = data["threat_score"]
        assert isinstance(score, int)
        assert 0 <= score <= 100

    def test_known_malicious_is_bool(self, n8n_webhook_url):
        data = self._post(n8n_webhook_url, "8.8.8.8")
        assert isinstance(data["known_malicious"], bool)

    def test_benign_ip_has_low_score(self, n8n_webhook_url):
        """Google DNS should score under 20."""
        data = self._post(n8n_webhook_url, "8.8.8.8")
        assert data["threat_score"] < 20

    def test_summary_key_present(self, n8n_webhook_url):
        """summary may be empty (AI disabled) but the key must exist."""
        data = self._post(n8n_webhook_url, "8.8.8.8")
        assert "summary" in data

    def test_response_survives_validate_response(self, n8n_webhook_url):
        """Live n8n data must survive _validate_response without crashing."""
        raw = self._post(n8n_webhook_url, "1.1.1.1")
        validated = _validate_response(raw)
        assert isinstance(validated["threat_score"], int)

    def test_domain_target_is_handled(self, n8n_webhook_url):
        """Domains route through the URL branch of the workflow."""
        data = self._post(n8n_webhook_url, "google.com")
        assert data["status"] == "success"
        assert isinstance(data["threat_score"], int)

    def test_empty_target_does_not_crash(self, n8n_webhook_url):
        """Empty target must return a graceful error, not a 500."""
        resp = requests.post(n8n_webhook_url, json={"target": ""}, timeout=20)
        assert resp.status_code in (200, 400, 422)
