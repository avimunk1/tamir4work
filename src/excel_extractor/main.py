"""Entry point for Excel data extraction."""

import logging
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from excel_extractor.extractor import extract_from_folder

OUTPUT_COLUMNS = ["source file name", "smel-mosad", "total hours"]


def _get_project_root() -> Path:
    """Project root: exe directory when frozen, else project root when running as script."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent.parent


def ensure_folders(input_dir: Path, output_dir: Path) -> None:
    """Create input and output folders if they do not exist."""
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    logging.info("Input folder: %s", input_dir)
    logging.info("Output folder: %s", output_dir)


def main() -> None:
    """Run extraction on all Excel files in input folder and write consolidated output."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    project_root = _get_project_root()
    input_dir = project_root / "input"
    output_dir = project_root / "output"

    ensure_folders(input_dir, output_dir)

    results = extract_from_folder(input_dir)

    if not results:
        logging.error(
            "No Excel files (.xlsx) found in input folder.\n"
            "Please add .xlsx files to: %s",
            input_dir,
        )
        sys.exit(1)

    df = pd.DataFrame(results)
    df.columns = OUTPUT_COLUMNS

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_dir / f"extracted_data_{timestamp}.xlsx"
    df.to_excel(output_file, index=False, engine="openpyxl")

    logging.info("Wrote %d rows to %s", len(df), output_file)


if __name__ == "__main__":
    main()
