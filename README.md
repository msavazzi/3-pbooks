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
- Fully configurable via **command-line arguments**.

---

## Requirements

- Python 3.8+
- Required Python packages:
  ```bash
  pip install pandas tqdm ebooklib openpyxl
  ```

> `openpyxl` is needed to read `.xlsx` files.

---

## Usage

Run the script with optional command-line arguments:

```bash
python create_books.py [--excel EXCEL_FILE] [--output OUTPUT_DIR] [--log LOG_FILE] [--create-pbook]
```

### Arguments

| Option | Description | Default |
|--------|-------------|---------|
| `--excel` | Path to the Excel file | `books.xlsx` |
| `--output` | Output directory for EPUB/PBOOK files | `pbooks/` next to Excel file |
| `--log` | Path to the log file | `pbook_creation_log.txt` next to Excel file |
| `--create-pbook` | Include this flag to also create `.pbook` files | Disabled by default |

### Examples

- Default run:
  ```bash
  python create_books.py
  ```

- Custom Excel and output folder:
  ```bash
  python create_books.py --excel mydata.xlsx --output output_books
  ```

- Save log to a custom path:
  ```bash
  python create_books.py --log /path/to/mylog.txt
  ```

- Also create `.pbook` files:
  ```bash
  python create_books.py --create-pbook
  ```

---

## Example

Excel row:

| Title | Author | Publication Date | ISBN | Source | Language |
|-------|--------|-----------------|------|--------|---------|
| My Book | John Doe | 2025-01-01 | 1234567890 | LocalLibrary | en |

Generated files:

- `John Doe - My Book.epub`
- `John Doe - My Book.pbook` *(if `--create-pbook` is used)*

EPUB metadata:

- Title: `My Book`
- Author: `John Doe`
- ISBN: `1234567890`
- Source: `LocalLibrary`
- Publication Date: `2025-01-01`
- Language: `en`
- Tag/Subject: `Physical`

---

## Notes

- Filenames are sanitized for Windows compatibility (invalid characters replaced).  
- If filenames are too long, they are intelligently truncated.  
- Duplicate filenames are automatically made unique.

---

## License

MIT License
