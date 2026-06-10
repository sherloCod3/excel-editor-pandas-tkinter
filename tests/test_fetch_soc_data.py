"""Unit tests for core.fetch_soc_data — mock mode and live mode (mocked requests)."""

import pytest
import requests
from unittest.mock import MagicMock, patch
from core import fetch_soc_data


# ─── Mock mode (no webhook URL) ───────────────────────────────────────────────
class TestMockMode:
    @pytest.mark.unit
    def test_returns_dict(self):
        result = fetch_soc_data("8.8.8.8")
        assert isinstance(result, dict)

    @pytest.mark.unit
    def test_has_all_required_keys(self):
        result = fetch_soc_data("1.2.3.4")
        for key in ("status", "target", "threat_score", "location", "known_malicious"):
            assert key in result, f"Missing key: {key}"

    @pytest.mark.unit
    def test_target_is_echoed(self):
        result = fetch_soc_data("192.168.1.50")
        assert result["target"] == "192.168.1.50"

    @pytest.mark.unit
    def test_threat_score_is_int_in_range(self):
        for _ in range(10):   # run several times given random output
            result = fetch_soc_data("10.0.0.1")
            assert isinstance(result["threat_score"], int)
            assert 0 <= result["threat_score"] <= 100

    @pytest.mark.unit
    def test_no_error_key_in_mock(self):
        result = fetch_soc_data("example.com", webhook_url=None)
        assert "error" not in result

    @pytest.mark.unit
    def test_empty_webhook_url_triggers_mock(self):
        result = fetch_soc_data("1.1.1.1", webhook_url="")
        assert result["status"] == "success"


# ─── Live mode — successful response ─────────────────────────────────────────
class TestLiveModeSuccess:
    @pytest.mark.unit
    def test_calls_webhook_with_target(self, monkeypatch):
        monkeypatch.setenv("SECOPS_ALLOW_LOCAL_WEBHOOKS", "1")
        mock_resp = MagicMock()
        mock_resp.json.return_value = {
            "status": "success",
            "target": "8.8.8.8",
            "threat_score": 5,
            "location": "United States",
            "known_malicious": False,
            "summary": "Clean IP.",
            "details": "ISP: Google LLC",
        }
        mock_resp.raise_for_status = MagicMock()

        with patch("core.requests.post", return_value=mock_resp) as mock_post:
            result = fetch_soc_data("8.8.8.8", webhook_url="http://localhost:5678/webhook/soc-scan")

        mock_post.assert_called_once_with(
            "http://localhost:5678/webhook/soc-scan",
            json={"target": "8.8.8.8"},
            timeout=15,
        )
        assert result["threat_score"] == 5
        assert result["location"] == "United States"

    @pytest.mark.unit
    def test_summary_field_passed_through(self, monkeypatch):
        monkeypatch.setenv("SECOPS_ALLOW_LOCAL_WEBHOOKS", "1")
        mock_resp = MagicMock()
        mock_resp.json.return_value = {
            "status": "success",
            "target": "8.8.8.8",
            "threat_score": 0,
            "location": "US",
            "known_malicious": False,
            "summary": "AI says: all clear.",
        }
        mock_resp.raise_for_status = MagicMock()

        with patch("core.requests.post", return_value=mock_resp):
            result = fetch_soc_data("8.8.8.8", webhook_url="http://localhost:5678/webhook/soc-scan")

        assert result["summary"] == "AI says: all clear."

    @pytest.mark.unit
    def test_string_score_from_n8n_is_coerced(self, monkeypatch):
        monkeypatch.setenv("SECOPS_ALLOW_LOCAL_WEBHOOKS", "1")
        """n8n might return threat_score as a string — must be cast to int."""
        mock_resp = MagicMock()
        mock_resp.json.return_value = {
            "status": "success",
            "target": "1.1.1.1",
            "threat_score": "72",   # ← string from n8n
            "location": "AU",
            "known_malicious": True,
        }
        mock_resp.raise_for_status = MagicMock()

        with patch("core.requests.post", return_value=mock_resp):
            result = fetch_soc_data("1.1.1.1", webhook_url="http://localhost:5678/webhook/soc-scan")

        assert result["threat_score"] == 72
        assert isinstance(result["threat_score"], int)


# ─── Live mode — error handling ───────────────────────────────────────────────
class TestLiveModeErrors:
    @pytest.mark.unit
    def test_connection_error_returns_error_key(self, monkeypatch):
        monkeypatch.setenv("SECOPS_ALLOW_LOCAL_WEBHOOKS", "1")
        with patch("core.requests.post", side_effect=requests.exceptions.ConnectionError("refused")):
            result = fetch_soc_data("8.8.8.8", webhook_url="http://localhost:5678/webhook/soc-scan")
        assert "error" in result
        assert "refused" in result["error"]

    @pytest.mark.unit
    def test_timeout_returns_error_key(self, monkeypatch):
        monkeypatch.setenv("SECOPS_ALLOW_LOCAL_WEBHOOKS", "1")
        with patch("core.requests.post", side_effect=requests.exceptions.Timeout("timed out")):
            result = fetch_soc_data("8.8.8.8", webhook_url="http://localhost:5678/webhook/soc-scan")
        assert "error" in result

    @pytest.mark.unit
    def test_http_error_returns_error_key(self, monkeypatch):
        monkeypatch.setenv("SECOPS_ALLOW_LOCAL_WEBHOOKS", "1")
        mock_resp = MagicMock()
        mock_resp.raise_for_status.side_effect = requests.exceptions.HTTPError("403 Forbidden")
        with patch("core.requests.post", return_value=mock_resp):
            result = fetch_soc_data("8.8.8.8", webhook_url="http://localhost:5678/webhook/soc-scan")
        assert "error" in result

    @pytest.mark.unit
    def test_uses_15s_timeout(self, monkeypatch):
        monkeypatch.setenv("SECOPS_ALLOW_LOCAL_WEBHOOKS", "1")
        """Ensures we don't accidentally revert to the old 10s timeout."""
        mock_resp = MagicMock()
        mock_resp.json.return_value = {"status": "success", "target": "x", "threat_score": 0,
                                       "location": "US", "known_malicious": False}
        mock_resp.raise_for_status = MagicMock()

        with patch("core.requests.post", return_value=mock_resp) as mock_post:
            fetch_soc_data("x", webhook_url="http://localhost:5678/webhook/soc-scan")

        _, kwargs = mock_post.call_args
        assert kwargs.get("timeout") == 15, "Timeout must be 15s to handle VirusTotal latency"
