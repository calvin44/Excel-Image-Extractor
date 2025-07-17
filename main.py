import os
import argparse
import json
from sys import stdout

from extractor import extract_images_from_excel

# Ensure stdout is set to UTF-8 encoding for proper output
stdout.reconfigure(encoding='utf-8')


def main():
    """
    Command-line interface for extracting images from an Excel file.

    Usage:
        python extract_images.py <excel_file> <sheet_name> <output_folder>

    On success, prints:
    {
        "success": true,
        "sheet_name": "Sheet1",
        "output_folder": "/abs/path/to/folder",
        "image_count": 3,
        "results": [...]
    }

    On error, prints:
    {
        "success": false,
        "error": "Detailed error message"
    }
    """
    parser = argparse.ArgumentParser(
        description="Extract images from Excel sheet and save to folder.")
    parser.add_argument("excel_file", help="Path to the Excel (.xlsx) file")
    parser.add_argument("sheet_name", help="Sheet name to extract images from")
    parser.add_argument(
        "output_folder", help="Folder path to save extracted images")

    args = parser.parse_args()

    try:
        results = extract_images_from_excel(
            args.excel_file, args.sheet_name, args.output_folder)

        output = {
            "success": True,
            "sheet_name": args.sheet_name,
            "output_folder": os.path.abspath(args.output_folder),
            "image_count": len(results),
            "results": results
        }

        print(json.dumps(output, indent=2, ensure_ascii=False))

    except Exception as e:  # pylint: disable=broad-except
        print(json.dumps({
            "success": False,
            "error": str(e)
        }, ensure_ascii=False))


if __name__ == "__main__":
    main()
