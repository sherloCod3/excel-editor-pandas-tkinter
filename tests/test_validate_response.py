"""Unit tests for core._validate_response — the schema guard."""

import pytest
from core import _validate_response


# ─── Happy-path tests ─────────────────────────────────────────────────────────
class TestValidResponsePassThrough:
    @pytest.mark.unit
    def test_full_valid_response_unchanged(self, valid_vt_response):
        result = _validate_response(valid_vt_response)
        assert result["status"] == "success"
        assert result["target"] == "8.8.8.8"
        assert result["threat_score"] == 2
        assert result["location"] == "United States"
        assert result["known_malicious"] is False
        assert "Google" in result["summary"]

    @pytest.mark.unit
    def test_threat_score_stays_int(self, valid_vt_response):
        result = _validate_response(valid_vt_response)
        assert isinstance(result["threat_score"], int)


# ─── Type coercion ────────────────────────────────────────────────────────────
class TestThreatScoreCasting:
    @pytest.mark.unit
    def test_string_score_is_cast_to_int(self):
        result = _validate_response({"threat_score": "42"})
        assert result["threat_score"] == 42
        assert isinstance(result["threat_score"], int)

    @pytest.mark.unit
    def test_float_score_is_cast_to_int(self):
        result = _validate_response({"threat_score": 73.9})
        assert result["threat_score"] == 73

    @pytest.mark.unit
    def test_none_score_defaults_to_zero(self):
        result = _validate_response({"threat_score": None})
        assert result["threat_score"] == 0

    @pytest.mark.unit
    def test_garbage_score_defaults_to_zero(self):
        result = _validate_response({"threat_score": "not-a-number"})
        assert result["threat_score"] == 0

    @pytest.mark.unit
    def test_score_clamped_at_100(self):
        result = _validate_response({"threat_score": 999})
        assert result["threat_score"] == 100

    @pytest.mark.unit
    def test_negative_score_clamped_at_zero(self):
        result = _validate_response({"threat_score": -10})
        assert result["threat_score"] == 0


# ─── Missing-field defaults ───────────────────────────────────────────────────
class TestMissingFieldDefaults:
    @pytest.mark.unit
    def test_empty_dict_gets_all_defaults(self):
        result = _validate_response({})
        assert result["status"] == "success"
        assert result["target"] == ""
        assert result["threat_score"] == 0
        assert result["location"] == "Unknown"
        assert result["known_malicious"] is False
        assert result["summary"] == ""
        assert result["details"] == ""

    @pytest.mark.unit
    def test_partial_response_fills_missing_fields(self):
        result = _validate_response({"target": "1.2.3.4", "threat_score": 55})
        assert result["target"] == "1.2.3.4"
        assert result["threat_score"] == 55
        assert result["location"] == "Unknown"   # default applied
        assert result["summary"] == ""            # default applied

    @pytest.mark.unit
    def test_n8n_wins_over_defaults(self):
        """n8n data should always take precedence over defaults."""
        result = _validate_response({"location": "Brazil", "status": "success"})
        assert result["location"] == "Brazil"

    @pytest.mark.unit
    def test_summary_field_passed_through(self):
        result = _validate_response({"summary": "Looks suspicious. Block it."})
        assert result["summary"] == "Looks suspicious. Block it."

    @pytest.mark.unit
    def test_extra_fields_preserved(self):
        """Unknown fields from n8n should not be stripped."""
        result = _validate_response({"isp": "Google LLC", "threat_score": 5})
        assert result["isp"] == "Google LLC"
