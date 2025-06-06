# SRT to DOCX Converter

A Python tool that converts SRT (SubRip) subtitle files to formatted Microsoft Word documents (DOCX).

## Features

- Convert SRT subtitle files to DOCX format
- Preserve timing information
- Clean, readable formatting in the output document
- Command-line interface for easy usage

## Installation

Install the package using pip:

```bash
pip install srt-to-docx
```

Or install from source:

```bash
git clone https://github.com/jpagh/srt-to-docx.git
cd srt-to-docx
pip install .
```

## Usage

Convert an SRT file to DOCX:

```bash
srt-to-docx input.srt
```

Convert all SRT files in a directory (recursively) to DOCX:

```bash
srt-to-docx input_directory
```

## Requirements

- Python 3.9 or higher
- docxtpl (for DOCX template processing)
- srt2 (for SRT file parsing)

## License

This project is licensed under the MIT License.
