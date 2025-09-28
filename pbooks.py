import pandas as pd
import os
import re
import unicodedata
from tqdm import tqdm
from ebooklib import epub
import argparse

# ---------------- Utility Functions ----------------
RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10))
}
MAX_FILENAME_LENGTH = 100

def sanitize_filename(filename):
    filename = unicodedata.normalize('NFKD', filename).encode('ascii', 'ignore').decode('ascii')
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    filename = filename.rstrip(' .')
    if filename.split('.')[0].upper() in RESERVED_NAMES:
        filename = '_' + filename
    return filename

def truncate_filename(author, title, ext=".pbook"):
    author_sanitized = sanitize_filename(author)
    title_sanitized = sanitize_filename(title)
    max_len = MAX_FILENAME_LENGTH - len(ext) - 3
    truncated = False
    if len(author_sanitized) + len(title_sanitized) > max_len:
        truncated = True
        total_len = len(author_sanitized) + len(title_sanitized)
        author_len = max(1, int(len(author_sanitized) / total_len * max_len))
        title_len = max(1, max_len - author_len)
        author_sanitized = author_sanitized[:author_len]
        title_sanitized = title_sanitized[:title_len]
    return f"{author_sanitized} - {title_sanitized}{ext}", truncated

def make_unique(filepath):
    base, ext = os.path.splitext(filepath)
    counter = 1
    renamed = False
    unique_path = filepath
    while os.path.exists(unique_path):
        unique_path = f"{base} ({counter}){ext}"
        counter += 1
        renamed = True
    return unique_path, renamed

# ---------------- EPUB Creation ----------------
def create_epub_file(filepath_base, title, author, isbn, source, pub_date, language):
    book = epub.EpubBook()
    
    # Set EPUB metadata fields
    book.set_identifier(isbn or "UnknownISBN")
    book.set_title(title or "Unknown Title")
    book.set_language(language or "en")
    book.add_author(author or "Unknown Author")
    book.add_metadata('DC', 'source', source or "")
    book.add_metadata('DC', 'date', pub_date or "")
    book.add_metadata('DC', 'subject', 'Physical')
    
    # Visible chapter
    content_html = f"""<h1>{title}</h1>
<p><strong>Author:</strong> {author}</p>
<p><strong>ISBN:</strong> {isbn}</p>
<p><strong>Source:</strong> {source}</p>
<p><strong>Publication Date:</strong> {pub_date}</p>
<p><strong>Language:</strong> {language}</p>
<p><strong>Tag:</strong> Physical</p>
"""
    chapter = epub.EpubHtml(title="Metadata", file_name='metadata.xhtml', content=content_html)
    book.add_item(chapter)
    
    # Add navigation files
    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    
    # Optional CSS
    style = 'body { font-family: Arial, sans-serif; } h1 { color: darkblue; }'
    nav_css = epub.EpubItem(uid="style_nav", file_name="style/nav.css", media_type="text/css", content=style)
    book.add_item(nav_css)
    
    # Spine: defines reading order
    book.spine = ['nav', chapter]
    
    # Write EPUB
    epub.write_epub(filepath_base + ".epub", book, {})

# ---------------- Main ----------------
def main():
    parser = argparse.ArgumentParser(description="Convert Excel rows to EPUB (and optionally PBOOK).")
    parser.add_argument("--excel", default="books.xlsx", help="Path to the Excel file (default: books.xlsx)")
    parser.add_argument("--output", help="Output directory for generated files (default: 'pbooks' next to Excel)")
    parser.add_argument("--log", help="Path to the log file (default: 'pbook_creation_log.txt' next to Excel)")
    parser.add_argument("--create-pbook", action="store_true", help="Also create .pbook files")
    args = parser.parse_args()

    excel_file = args.excel
    base_dir = os.path.dirname(os.path.abspath(excel_file)) or "."

    output_dir = args.output or os.path.join(base_dir, "pbooks")
    os.makedirs(output_dir, exist_ok=True)

    log_file_path = args.log or os.path.join(base_dir, "pbook_creation_log.txt")
    CREATE_PBOOK = args.create_pbook

    # Read Excel
    df = pd.read_excel(excel_file)

    # Process rows
    with open(log_file_path, 'w', encoding='utf-8') as log_file:
        log_file.write("Filename\tOriginal Author\tOriginal Title\tTruncated\tRename\n")
        
        for row in tqdm(df.itertuples(index=False), total=len(df), desc="Creating files"):
            title = str(getattr(row, 'Title', '')).strip()
            author = str(getattr(row, 'Author', '')).strip()
            isbn = str(getattr(row, 'ISBN', '')).strip()
            source = str(getattr(row, 'Source', '')).strip()
            pub_date = str(getattr(row, 'Publication Date', '')).strip()
            language = str(getattr(row, 'Language', 'en')).strip()
            
            filename, truncated = truncate_filename(author, title)
            filepath_base = os.path.join(output_dir, os.path.splitext(filename)[0])
            filepath_base, renamed = make_unique(filepath_base)
            
            # Create .pbook if option enabled
            if CREATE_PBOOK:
                pbook_path = filepath_base + ".pbook"
                with open(pbook_path, 'w', encoding='utf-8') as f:
                    f.write("[metadata]\n")
                    f.write(f"title={title}\n")
                    f.write(f"author={author}\n")
                    f.write(f"ISBN={isbn}\n")
                    f.write(f"source={source}\n")
                    f.write(f"publication_date={pub_date}\n")
                    f.write(f"language={language}\n")
            else:
                pbook_path = filepath_base + ".pbook"  # for logging only
            
            # Create EPUB
            create_epub_file(filepath_base, title, author, isbn, source, pub_date, language)
            
            # Log entry
            log_file.write(f"{os.path.basename(pbook_path)}\t{author}\t{title}\t{truncated}\t{renamed}\n")

    print(f"All {len(df)} files created. Log saved to '{log_file_path}'.")

if __name__ == "__main__":
    main()
