# Email Extractor

A Python-based tool with a graphical interface (GUI) to extract email addresses from files and directories. The tool supports multiple file types and includes filters for specific email domains.

## Features
- **Extract Emails**: Extract email addresses from `.txt`, `.csv`, `.sql`, `.docx`, `.xlsx`, and `.zip` files.
- **Domain Filters**: Specify domain filters to extract emails from specific domains.
- **Save Results**: Save extracted email addresses to a `.txt` file.
- **Progress Tracking**: Real-time progress updates with elapsed time tracking.

## Requirements
- Python 3.8+
- Dependencies:
  - `pandas`
  - `python-docx`
  - `openpyxl`

Install dependencies using:
```bash
pip install pandas python-docx openpyxl
