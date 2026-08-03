# Editable PDF to Word

A minimal Python tool for converting text-based PDF files into editable DOCX
documents. It uses `pdf2docx` to reconstruct the layout, PyMuPDF to inspect and
validate the PDF, FontTools to inspect embedded fonts, and `python-docx` to fix
common image and heading issues after conversion.

## Features

- Converts PDF text, images, tables, and page breaks into an editable DOCX.
- Preserves text formatting and embeds permitted fonts found in the PDF.
- Fixes common image-spacing and bullet-font issues.
- Optionally assigns Word Heading styles and repairs short fallback-font runs.
- Runs locally without uploading files or downloading fonts.

## Requirements

- Python 3.10 or later; Python 3.12 is recommended.
- A text-based PDF. Scanned PDFs require an OCR step before conversion.

## Installation

Run the commands once from the project directory.

Linux, macOS, or WSL:

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install --upgrade pip && python3 -m pip install -r requirements.txt
```

Windows PowerShell:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip; if ($LASTEXITCODE -eq 0) { python -m pip install -r requirements.txt }
```

## Usage

Keep the virtual environment active and run:

```bash
python3 main.py input.pdf
```

Choose the output file:

```bash
python3 main.py input.pdf -o output.docx
```

Input and output files may be anywhere on the filesystem. Quote paths that
contain spaces:

```bash
python3 main.py "/path/to/input file.pdf" -o "/path/to/output file.docx"
```

Detect numbered headings and assign native Word Heading styles:

```bash
python3 main.py input.pdf -o output.docx --headings
```

Overwrite an existing output file:

```bash
python3 main.py input.pdf -o output.docx --force
```

Keep font names but create a smaller DOCX without embedded PDF fonts:

```bash
python3 main.py input.pdf -o output.docx --no-embed-fonts
```

On native Windows, replace `python3` in the usage examples with `python` after
activating the virtual environment.

## Font handling

- Font families are recovered from PDF metadata without hard-coded Calibri or
  Arial replacements. Size, weight, italic, color, script-specific fonts, and
  Heading formatting are preserved when available.
- Embedded TrueType/OpenType fonts are copied from the PDF only when their
  license flags permit it. Missing, restricted, or invalid fonts produce a
  `Font warning`; the tool never searches for or downloads fonts.
- PDF fonts are often subsets. If a normalized heading needs missing glyphs,
  the incomplete subset is skipped and Word uses the installed full font.
- Use `--no-embed-fonts` to keep font names without including font files in the
  DOCX. Word may substitute fonts that are not installed.

## Troubleshooting

If Python reports `No module named ...`, install the dependencies with the same
Python interpreter used to run the tool.

With the project's virtual environment:

```bash
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

Without a virtual environment:

```bash
python3 -m pip install --user -r requirements.txt
```

Use a `.docx` extension when setting the output path.

## Limitations

- DOCX reflows content, so page layout and detected headings **should be reviewed**.
- Scanned PDFs require OCR; complex equations, columns, and tables may need
  **manual correction**.
- Fonts cannot be embedded when they are missing from the PDF or their license
  flags prohibit it.

## License

Released under the [MIT License](LICENSE).
