# Markdown to Word Converter

This Python script `md_to_docx.py` converts Markdown files to Word documents (`.docx`). It's designed to handle Markdown formatting, including headings, paragraphs, and code blocks, while ensuring proper display of Chinese fonts and correct formatting of code blocks.

## Features

-   Converts Markdown files to Word documents.
-   Supports headings (`h1` to `h3`), paragraphs, and code blocks.
-   Uses a text-to-image approach for rendering code blocks, preserving formatting.
-   Ensures proper display of Chinese characters in headings and document text.
-   Provides batch conversion for processing multiple Markdown files in a directory.

## Prerequisites

-   Python 3.9 or later
-   Required Python libraries:
    -   `markdown`
    -   `python-docx`
    -   `beautifulsoup4`
    -   `Pillow`

You can install these libraries using `pip`:

```bash
pip install -r requirements.txt
```
## Usage
### Single File Conversion
To convert a single Markdown file to a Word document, run the following command:
```bash
python md_to_docx.py your_file.md
```
