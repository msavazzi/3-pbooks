# Excel to EPUB / PBOOK Converter

This Python script reads an Excel file with book metadata and generates **EPUB files** for each entry. Optionally, it can also create `.pbook` files containing the same metadata. The EPUB files include proper metadata compatible with Calibre.

---

## Features

- Reads an Excel file with columns:
  - `Title`
  - `Author`
  - `Publisher` *(optional, not used in metadata but can be added)*
  - `Publication Date`
  - `ISBN`
  - `Source`
  - `Language` *(optional, defaults to `en`)*
- Creates **EPUB files** with:
  - Title, Author, ISBN, Source, Publication Date, Language
  - Tag/Subject: `Physical`
- Optional `.pbook` file creation.
- Windows-safe filenames with intelligent truncation and uniqueness handling.
- Tab-separated log file listing all generated files.
- Progress bar for monitoring conversion.

---

## Requirements

- Python 3.8+
- Required Python packages:
  ```bash
  pip install pandas tqdm ebooklib openpyxl
