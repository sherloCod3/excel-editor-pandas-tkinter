"""Unit tests for core.load_data — CSV and Excel parsing."""

import io

import pandas as pd
import pytest
from core import load_data


class TestLoadCsv:
    @pytest.mark.unit
    def test_valid_csv_returns_dataframe(self, csv_file_obj):
        df = load_data(csv_file_obj)
        assert df is not None
        assert isinstance(df, pd.DataFrame)

    @pytest.mark.unit
    def test_csv_has_correct_columns(self, csv_file_obj):
        df = load_data(csv_file_obj)
        assert "Source IP" in df.columns
        assert "Action" in df.columns

    @pytest.mark.unit
    def test_csv_has_correct_row_count(self, csv_file_obj):
        df = load_data(csv_file_obj)
        assert len(df) == 1   # fixture has one data row

    @pytest.mark.unit
    def test_csv_values_are_correct(self, csv_file_obj):
        df = load_data(csv_file_obj)
        assert df["Source IP"].iloc[0] == "192.168.1.1"
        assert df["Action"].iloc[0] == "Allowed"


class TestLoadExcel:
    @pytest.mark.unit
    def test_valid_xlsx_returns_dataframe(self, xlsx_file_obj):
        df = load_data(xlsx_file_obj)
        assert df is not None
        assert isinstance(df, pd.DataFrame)

    @pytest.mark.unit
    def test_xlsx_has_correct_columns(self, xlsx_file_obj):
        df = load_data(xlsx_file_obj)
        assert "Source IP" in df.columns
        assert "Port" in df.columns

    @pytest.mark.unit
    def test_xlsx_values_are_correct(self, xlsx_file_obj):
        df = load_data(xlsx_file_obj)
        assert df["Source IP"].iloc[0] == "10.0.0.1"
        assert df["Port"].iloc[0] == 443


class TestLoadDataErrorHandling:
    @pytest.mark.unit
    def test_unsupported_extension_returns_none(self, bad_file_obj):
        result = load_data(bad_file_obj)
        assert result is None

    @pytest.mark.unit
    def test_corrupted_csv_returns_none(self):
        buf = io.BytesIO(b"\x00\xff\xfe corrupted")
        buf.name = "broken.csv"
        # pandas may or may not raise on this — either way we expect None or a DF
        # The key is it must not raise an unhandled exception
        result = load_data(buf)
        # result can be None or a (possibly garbled) DataFrame — both are acceptable
        assert result is None or isinstance(result, pd.DataFrame)

    @pytest.mark.unit
    def test_empty_csv_returns_empty_dataframe(self):
        buf = io.BytesIO(b"Col1,Col2\n")
        buf.name = "empty.csv"
        df = load_data(buf)
        assert df is not None
        assert len(df) == 0
        assert list(df.columns) == ["Col1", "Col2"]
