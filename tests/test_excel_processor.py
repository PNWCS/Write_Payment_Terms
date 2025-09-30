"""Tests for Excel processing functions.

This module tests the core Excel processing functionality and payment terms import.
"""

import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

import pytest
from openpyxl import Workbook

from xlsx_reader.excel_processor import (
    PaymentTerm,
    create_payment_terms_batch_qbxml,
    process_payment_terms,
    read_payment_terms,
    save_payment_terms_to_quickbooks,
)


def create_payment_terms_excel(file_path: str) -> None:
    """Create a test Excel file with payment terms data."""
    workbook = Workbook()

    # Remove default sheet
    workbook.remove(workbook.active)

    # Create payment_terms sheet
    sheet = workbook.create_sheet("payment_terms")
    sheet["A1"] = "Name"
    sheet["B1"] = "ID"

    # Add test payment terms
    payment_terms_data = [
        ("Net 30", 30),
        ("Net 15", 15),
        ("Net 60", 60),
        ("2/10 Net 30", 10),
        ("Cash On Delivery", 0),
    ]

    for i, (name, discount_days) in enumerate(payment_terms_data, start=2):
        sheet[f"A{i}"] = name
        sheet[f"B{i}"] = discount_days

    workbook.save(file_path)


class TestPaymentTerms:
    """Test cases for payment terms functionality."""

    @pytest.fixture
    def payment_terms_excel_file(self):
        """Create a temporary Excel file with payment terms for testing."""
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_path = Path(tmp.name)
        try:
            tmp.close()
            create_payment_terms_excel(str(tmp_path))
            yield str(tmp_path)
        finally:
            try:
                if tmp_path.exists():
                    tmp_path.unlink()
            except PermissionError:
                pass

    def test_payment_term_dataclass(self):
        """Test PaymentTerm dataclass."""
        term = PaymentTerm(name="Net 30", discount_days=30)
        assert term.name == "Net 30"
        assert term.discount_days == 30

    def test_read_payment_terms(self, payment_terms_excel_file):
        """Test reading payment terms from Excel file."""
        payment_terms = read_payment_terms(payment_terms_excel_file)

        assert len(payment_terms) == 5
        assert payment_terms[0].name == "Net 30"
        assert payment_terms[0].discount_days == 30
        assert payment_terms[1].name == "Net 15"
        assert payment_terms[1].discount_days == 15
        assert payment_terms[4].name == "Cash On Delivery"
        assert payment_terms[4].discount_days == 0

    def test_create_payment_terms_batch_qbxml(self):
        """Test batch QBXML generation for payment terms."""
        terms = [
            PaymentTerm(name="Net 30", discount_days=30),
            PaymentTerm(name="Net 15", discount_days=15),
        ]
        qbxml = create_payment_terms_batch_qbxml(terms)
        assert "<?xml version=" in qbxml
        assert "<StandardTermsAdd>" in qbxml
        assert "<Name>Net 30</Name>" in qbxml
        assert "<StdDueDays >30</StdDueDays >" in qbxml
        assert "<Name>Net 15</Name>" in qbxml
        assert "<StdDueDays >15</StdDueDays >" in qbxml

    @patch("xlsx_reader.excel_processor.win32com.client.Dispatch")
    def test_save_payment_terms_to_quickbooks_success(self, mock_dispatch):
        """Test successful save to QuickBooks."""
        # Mock the COM objects
        mock_qb_app = Mock()
        mock_session = "test_session"
        mock_qb_app.BeginSession.return_value = mock_session
        mock_qb_app.ProcessRequest.return_value = '<?xml version="1.0"?><QBXML><QBXMLMsgsRs><StandardTermsAddRs statusCode="0" statusSeverity="Info"><StandardTermsRet><Name>Net 30</Name></StandardTermsRet></StandardTermsAddRs><StandardTermsAddRs statusCode="0" statusSeverity="Info"><StandardTermsRet><Name>Net 15</Name></StandardTermsRet></StandardTermsAddRs></QBXMLMsgsRs></QBXML>'
        mock_dispatch.return_value = mock_qb_app

        payment_terms = [
            PaymentTerm(name="Net 30", discount_days=30),
            PaymentTerm(name="Net 15", discount_days=15),
        ]

        result = save_payment_terms_to_quickbooks(payment_terms)

        assert len(result) == 2
        assert "Net 30" in result
        assert "Net 15" in result
        mock_qb_app.OpenConnection.assert_called_once()
        mock_qb_app.BeginSession.assert_called_once()
        mock_qb_app.EndSession.assert_called_once()
        mock_qb_app.CloseConnection.assert_called_once()

    @patch("xlsx_reader.excel_processor.win32com.client.Dispatch")
    def test_save_payment_terms_to_quickbooks_connection_error(self, mock_dispatch):
        """Test handling of QuickBooks connection error."""
        mock_dispatch.side_effect = Exception("QuickBooks not running")

        payment_terms = [PaymentTerm(name="Net 30", discount_days=30)]

        with pytest.raises(RuntimeError, match="Failed to connect to QuickBooks"):
            save_payment_terms_to_quickbooks(payment_terms)

    def test_read_payment_terms_file_not_found(self):
        """Test handling of non-existent payment terms file."""
        with pytest.raises(FileNotFoundError):
            read_payment_terms("nonexistent.xlsx")

    @patch("xlsx_reader.excel_processor.save_payment_terms_to_quickbooks")
    def test_process_payment_terms(self, mock_save, payment_terms_excel_file):
        """Test the complete payment terms processing workflow."""
        mock_save.return_value = ["Net 30", "Net 15"]

        result = process_payment_terms(payment_terms_excel_file)

        assert result == ["Net 30", "Net 15"]
        mock_save.assert_called_once()
