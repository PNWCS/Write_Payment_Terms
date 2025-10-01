"""Excel processing module for reading XLSX files and QuickBooks integration.

This module provides functions to read Excel files, specifically payment terms,
and integrate with QuickBooks Desktop via COM API.
"""

from dataclasses import dataclass
from typing import Any

import win32com.client
from openpyxl import load_workbook

import xml.etree.ElementTree as ET

@dataclass
class PaymentTerm:
    """Represents a payment term with name and discount days."""

    name: str
    discount_days: int


def read_payment_terms(file_path: str) -> list[PaymentTerm]:
    """Read payment terms from the specified Excel file.

    Expected Excel format:
    - Must contain a sheet named 'payment_terms'
    - Column A: Payment term names (strings)
    - Column B: Discount days (integers)
    - Row 1 should contain headers (will be skipped)
    - Data starts from row 2

    Args:
        file_path (str): Path to the Excel file containing payment terms (.xlsx format)

    Returns:
        list[PaymentTerm]: List of payment terms with name and discount days.
                          Empty list if no valid payment terms found.

    Raises:
        No exceptions need to be manually raised - let openpyxl handle file/sheet errors

    Implementation Notes:
        - Use openpyxl.load_workbook() with read_only=True for better performance
        - Access the 'payment_terms' worksheet by name
        - Use worksheet.iter_rows(min_row=2, values_only=True) to skip headers
        - Validate that both name (column A) and discount_days (column B) are not None
        - Convert name to string and strip whitespace
        - Convert discount_days to integer, skip rows with invalid data
        - Handle ValueError/TypeError when converting discount_days to int
    """
    payment_terms_list: list[PaymentTerm] = []
    sheetname = "payment_terms"

    # 1. Load the workbook (File Not Found handled by openpyxl/Python)
    # Use read_only=True for performance.
    workbook = load_workbook(file_path, read_only=True)

    # 2. Access the worksheet (Missing sheet handled by KeyError)
    worksheet = workbook[sheetname]

    # 3. Iterate over rows, skipping the header (min_row=2)
    # max_col=2 ensures we only read columns A (name) and B (discount_days)
    for row in worksheet.iter_rows(min_row=2, max_col=2, values_only=True):
        name_raw, discount_days_raw = row[0], row[1]

        # Validate that both column A (name) and B (discount_days) are not None
        if name_raw is None or discount_days_raw is None:
            continue

        try:
            # Convert name to string and strip whitespace
            name = str(name_raw).strip()

            # Skip if name is empty after stripping (e.g., if raw was just spaces)
            if not name:
                continue

            # Convert discount_days to integer
            discount_days = int(discount_days_raw)

            # Create and append the dataclass instance
            payment_terms_list.append(PaymentTerm(name=name, discount_days=discount_days))

        # Skip rows where discount_days conversion fails (e.g., 'TEN')
        except (ValueError, TypeError):
            continue

    return payment_terms_list


def connect_to_quickbooks() -> Any:
    """Connect to QuickBooks Desktop via COM API.

    This function establishes a connection to QuickBooks Desktop using the
    QBXML Request Processor COM interface. QuickBooks Desktop must be running
    with a company file open.

    Returns:
        tuple[Any, Any]: A tuple containing (qb_app, session)
            - qb_app: COM object for QuickBooks application interface
            - session: Session ticket for the current QB connection

    Raises:
        No exceptions need to be manually raised - let win32com handle COM errors

    Security Notes:
        - User may need to grant permission in QuickBooks for first-time access
        - QuickBooks may prompt user to allow external application access
    """
    try:
        qb_app = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        qb_app.OpenConnection("", "Payment Terms Import")
        session = qb_app.BeginSession("", 2)  # 2 = qbFileOpenDoNotCare
        return qb_app, session
    except Exception as e:
        print(f"QuickBooks connection error: {str(e)}")
        raise


def create_payment_terms_batch_qbxml(payment_terms: list[PaymentTerm]) -> str:
    """Create QBXML for adding multiple payment terms in a batch.

    This function generates a well-formed QBXML document containing multiple
    StandardTermsAddRq requests that can be sent to QuickBooks Desktop in a
    single batch operation.

    Args:
        payment_terms (list[PaymentTerm]): List of payment terms to create.
                                         Each PaymentTerm must have name and discount_days.

    Returns:
        str: Complete QBXML string ready to send to QuickBooks Desktop.
             Contains XML declaration, QBXML root, and multiple StandardTermsAddRq elements.

    Raises:
        AttributeError: If PaymentTerm objects are missing required attributes
        TypeError: If payment_terms is not a list or contains invalid objects

    QBXML Structure:
        <?xml version="1.0" encoding="utf-8"?>
        <?qbxml version="13.0"?>
        <QBXML>
            <QBXMLMsgsRq onError="continueOnError">
                <StandardTermsAddRq>
                    <StandardTermsAdd>
                        <Name>Term Name</Name>
                        <StdDueDays>Number of Days</StdDueDays>
                    </StandardTermsAdd>
                </StandardTermsAddRq>
                ... (repeated for each payment term)
            </QBXMLMsgsRq>
        </QBXML>

    Implementation Notes:
        - Create a list to store individual StandardTermsAddRq XML strings
        - Loop through payment_terms and create XML for each term
        - Use f-strings to format XML with term.name and term.discount_days
        - XML escape special characters in term names if necessary
        - Join all requests with newlines using chr(10).join()
        - Wrap in proper QBXML envelope with version="13.0"
        - Use onError="continueOnError" to process all terms even if some fail
        - Note: <StdDueDays > has trailing space - this is required by QB format
    """
    xml_strings = []
    for term in payment_terms:
        if not isinstance(term, PaymentTerm):
            raise TypeError("payment_terms must be a list of PaymentTerm objects")
        if not hasattr(term, "name") or not hasattr(term, "discount_days"):
            raise AttributeError("Each PaymentTerm must have 'name' and 'discount_days' attributes")
        if term.name is None or term.discount_days is None:
            raise ValueError("PaymentTerm 'name' and 'discount_days' cannot be None")
        if not isinstance(term.discount_days, int):
            raise TypeError("PaymentTerm 'discount_days' must be an integer")
        if not isinstance(term.name, str):
            raise TypeError("PaymentTerm 'name' must be a string")
        if term.name.strip() == "":
            raise ValueError("PaymentTerm 'name' cannot be empty or whitespace")
        # Ensure no special characters in name that could break XML
        term.name = term.name.strip().replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        # Append the XML for this term to the list
    
        term_xml = f"""
        <StandardTermsAddRq>
            <StandardTermsAdd>
                <Name>{term.name}</Name>
                <StdDueDays >{term.discount_days}</StdDueDays >
            </StandardTermsAdd> 
        </StandardTermsAddRq>"""
        xml_strings.append(term_xml)
        full_qbxml = f""" 
        <?xml version="1.0" encoding="utf-8"?>
            <?qbxml version="13.0"?>
            <QBXML>
                <QBXMLMsgsRq onError="continueOnError"> {chr(10).join(xml_strings)} </QBXMLMsgsRq>
            </QBXML>"""
    return full_qbxml

def save_payment_terms_to_quickbooks(payment_terms: list[PaymentTerm]) -> list[str]:
    """Save payment terms to QuickBooks Desktop.

    This function connects to QuickBooks, sends a batch QBXML request to create
    multiple payment terms, parses the response, and returns the list of
    successfully created terms.

    Args:
        payment_terms (list[PaymentTerm]): List of payment terms to save to QuickBooks.
                                         Each term must have valid name and discount_days.

    Returns:
        list[str]: List of payment term names that were successfully created.
                  May be shorter than input list if some terms failed or already exist.

    Raises:
        RuntimeError: If connection to QuickBooks fails (manually wrap exceptions for clarity)

    Response Parsing Logic:
        - Success: statusCode="0" indicates successful creation
        - Already exists: Error code "3100" means term already exists (skip silently)
        - Other errors: Print warning and exclude from results

    Implementation Notes:
        - Call connect_to_quickbooks() to establish QB connection
        - Use create_payment_terms_batch_qbxml() to generate QBXML request
        - Send QBXML using qb_app.ProcessRequest(session, qbxml)
        - Parse XML response to identify successful operations:
            * Look for statusCode="0" AND term name in response
            * Handle "3100" error code (term already exists) gracefully
            * Print warnings for other failures
        - Always cleanup: call qb_app.EndSession(session) and qb_app.CloseConnection()
        - Use try/except to catch and re-raise exceptions with meaningful messages
        - Consider that QB response may contain multiple StandardTermsAddRs elements

    QuickBooks Error Codes:
        - 0: Success
        - 3100: Object already exists
        - Other codes indicate various QB-specific errors
    """
    try:
        qb_app, session = connect_to_quickbooks()
    except Exception as e:
        raise RuntimeError("Failed to connect to QuickBooks") from e

    qbxml_request = create_payment_terms_batch_qbxml(payment_terms)
    try:
        qbxml_response = qb_app.ProcessRequest(session, qbxml_request)
    except Exception as e:
        raise RuntimeError("Failed to process request in QuickBooks") from e
    finally:
        try:
            qb_app.EndSession(session)
            qb_app.CloseConnection()
        except Exception:
            pass  # Ignore errors during cleanup

    created_terms = []
    try:
        root = ET.fromstring(qbxml_response)
        for add_rs in root.findall(".//StandardTermsAddRs"):
            status_code = add_rs.get("statusCode")
            if status_code == "0":
                term_name = add_rs.find(".//Name").text
                created_terms.append(term_name)
            elif status_code == "3100":
                # Term already exists, skip silently
                continue
            else:
                error_message = add_rs.get("statusMessage", "Unknown error")
                print(f"Warning: Failed to add term - {error_message}")
    except ET.ParseError as e:
        raise RuntimeError("Failed to parse QuickBooks response") from e

    return created_terms


def process_payment_terms(file_path: str) -> list[str]:
    """Read payment terms from Excel and save to QuickBooks.

    This is the main orchestration function that handles the complete workflow:
    reading payment terms from an Excel file and saving them to QuickBooks Desktop.

    Args:
        file_path (str): Path to the Excel file containing payment terms (.xlsx format).
                        File must contain a 'payment_terms' sheet with Name and ID columns.

    Returns:
        list[str]: List of payment term names that were successfully created in QuickBooks.
                  May be empty if no terms were processed or all failed.

    Raises:
        ValueError: If no payment terms found in the Excel file (manually check and raise)

    Workflow:
        1. Read payment terms from Excel file using read_payment_terms()
        2. Validate that at least one payment term was found
        3. Print summary of terms to be imported (for user feedback)
        4. Save all terms to QuickBooks using save_payment_terms_to_quickbooks()

    Implementation Notes:
        - Call read_payment_terms(file_path) to extract payment terms from Excel
        - Check if payment_terms list is empty and raise ValueError with helpful message
        - Print found terms for user visibility before QB operation
        - Use f"Found {len(payment_terms)} payment terms to import:" for count
        - Print each term as f"  - {term.name} ({term.discount_days} days)"
        - Let QuickBooks connection errors bubble up naturally from save function
        - No need for separate QB connection test - let save_payment_terms handle it

    Error Handling Strategy:
        - Validate Excel data before attempting QuickBooks operations
        - Provide clear error messages for common issues
        - Let underlying functions handle their specific error cases
        - Don't catch and re-wrap exceptions unless adding meaningful context
    """
    payment_terms = read_payment_terms(file_path)
    if not payment_terms:
        raise ValueError("No payment terms found in the Excel file.")
    
    print(f"Found {len(payment_terms)} payment terms to import:")
    for term in payment_terms:
        print(f"  - {term.name} ({term.discount_days} days)")
    
    created_terms = save_payment_terms_to_quickbooks(payment_terms)
    return created_terms
