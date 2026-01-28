from docx import Document
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
from src.config import StyleConfig
from src.header import CoverPageProcessor
from src.table import TableProcessor
from src.validator import validate_output
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

def process_document(input_path: Path,
                     output_path: Path,
                     validate: bool = True,
                     save: bool = True) -> None:
    """
    Core logic to apply styles, headers, and table formatting to a docx file.
    If validation fails, saves the file with a '_WITH_ISSUES' suffix and writes
    a log file.
    """
    
    # Check if file exists before determining to load it
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    logger.info("Loading document: %s", input_path.name)
    
    # Load Document
    doc = Document(str(input_path))

    # Global Normalization
    if 'Normal' in doc.styles:
        style = doc.styles['Normal']
        font = style.font
        font.name = StyleConfig.FONT_NAME
        font.size = StyleConfig.FONT_SIZE
        logger.debug("Applied global font settings: %s", StyleConfig.FONT_NAME)
    else:
        logger.warning("Style 'Normal' missing. Skipping font normalization.")

    # Process Cover Page
    logger.info("Processing cover page...")
    cover_processor = CoverPageProcessor(doc)
    cover_processor.process()

    # Process Tables
    logger.info("Processing tables...")
    table_processor = TableProcessor(doc)
    table_processor.process()

    # Ensure the output directory exists
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Optional: Validate the output
    issues = []
    if validate:
        issues = validate_output(doc)
    
    if issues:
        if save:
            base_name = output_path.stem
            issue_doc_path = output_path.parent / f"{base_name}_WITH_ISSUES.docx"
            issue_log_path = output_path.parent / f"{base_name}_ISSUES.txt"
            
            # Save the "bad" document so we can inspect it
            doc.save(str(issue_doc_path))

            # Save the issues list to a text file
            with open(issue_log_path, "w") as f:
                f.write(f"Validation Report for {input_path.name}\n")
                f.write("=" * 40 + "\n\n")
                for issue in issues:
                    f.write(f"- {issue}\n")

            logger.error(f"Validation FAILED. Output saved to: {issue_doc_path}")
            logger.error(f"Issue log saved to: {issue_log_path}")
            
            # Return early so that the program doesn't save the document
            return None, []
        
        else:
            return doc, issues
    else:
        if save:
            # Save normally if everything passed
            doc.save(str(output_path))
            logger.info("Processing complete. Validation PASSED. Saved to: %s", output_path)
            return None, []
        else:
            logger.info("Processing complete. Returning document object.")
            return doc, []
    
    