import argparse
import logging
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
from pathlib import Path


# Safety check for importing main functions
try:
    from src.logger import setup_logging
    from src.processor import process_document
except ImportError as e:
    sys.exit(f"Setup Error: Could not import modules from 'src'. \nDetails: {e}")

# Initialize logger variable
logger = logging.getLogger("CLI")

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Docx Formatter: Automate style, header, and table formatting."
    )
    
    parser.add_argument(
        '-i', '--input',
        type=Path,
        required=True,
        help="Path to the input .docx file."
    )
    
    parser.add_argument(
        '-o', '--output',
        type=Path,
        required=False,
        help="Path to save the processed file."
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help="Enable verbose debug logging."
    )

    parser.add_argument(
        '--validate',
        action='store_true',
        help="Run post-processing validation. If formatting issues are found, the file will be saved with a '_WITH_ISSUES' suffix."
    )

    return parser.parse_args()

def main():
    args = parse_args()

    # Configure Logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    setup_logging(level=log_level)
    
    # Now we can use the logger
    logger.info("Application started.")

    input_path: Path = args.input.resolve()
    
    if args.output:
        output_path: Path = args.output.resolve()
    else:
        output_path = input_path.with_name(f"{input_path.stem}_processed{input_path.suffix}")

    try:
        process_document(input_path, output_path, validate=args.validate)
    except Exception as e:
        logger.error("Couldn't process the document: %s", e)
        if args.verbose:
            logger.exception("Full traceback:")
        sys.exit(1)

if __name__ == "__main__":
    main()