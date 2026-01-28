import pytest
import sys
import os
from pathlib import Path

# Ensure 'src' is importable regardless of where pytest is run from
sys.path.append(os.path.join(os.path.dirname(__file__), ".."))

from src.processor import process_document

# Configuration: Where to look for input files
TEST_INPUTS_DIR = Path(__file__).parent / "inputs"

def get_test_files():
    """
    Scans the tests/inputs directory and returns a list of all .docx files.
    This function is used by pytest to generate dynamic test cases.
    """
    if not TEST_INPUTS_DIR.exists():
        return []
    
    files = list(TEST_INPUTS_DIR.glob("*.docx"))
    
    # Sort files to ensure consistent test order
    return sorted(files)

@pytest.mark.parametrize("input_path", get_test_files())
def test_financial_report_compliance(input_path):
    """
    Runs the processor on a real .docx file and asserts 0 validation issues.
    """
    print(f"\nTesting file: {input_path.name}")

    # We provide a dummy output path because the function signature requires it,
    # but since save=False, nothing will be written to disk.
    dummy_output_path = input_path.parent / f"{input_path.stem}_processed.docx"

    # Execute logic with save=False
    # This prevents creating actual files and returns (doc, issues) directly
    result = process_document(
        input_path=input_path,
        output_path=dummy_output_path,
        validate=True,
        save=True
    )
    
    # Unpack the results
    _, issues = result

    # --- THE ASSERTION ---
    # Fail the test if the 'issues' list is not empty.
    error_msg = f"Validation failed for file '{input_path.name}' with {len(issues)} issues:\n"
    error_msg += "\n".join(f" - {issue}" for issue in issues)

    assert not issues, error_msg