# Excel Image Extractor

A Python command-line utility that extracts all images from a specified sheet in an Excel `.xlsx` file. Each image is saved to a given folder along with metadata including the cell it is anchored to and its (row, col) position.

## 📦 Features

- Extracts all embedded images from a given sheet
- Returns:
  - Image file path
  - Excel cell anchor (e.g., `B3`)
  - Column and row numbers (1-based)
- Clears the output folder before saving new images
- Can be packaged as a `.exe` using PyInstaller

---

## 🔧 Requirements

- Python 3.7+
- `openpyxl`
- `pillow` (implicitly required by `openpyxl` for image processing)

Install dependencies:

```bash
pip install openpyxl pillow
```

 
## 🚀 Usage (CLI)
```bash
python main.py <excel_file> <sheet_name> <output_folder>
```


# This will:
- Clear the output_images/ folder
- Save extracted images into it
- Print the metadata to the console