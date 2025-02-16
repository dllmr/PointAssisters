# PointAssisters - PowerPoint Analyzer

A command-line tool and a corresponding GUI tool (with a simple Qt UI) to analyze PowerPoint presentations (.pptx files) and report on:
- Hidden slides
- Slides featuring animations and transitions
- Font usage and missing fonts

## Features

- **Hidden Slides Detection**: Identifies any slides marked as hidden in the presentation
- **Effects Analysis**: Lists slides containing animations or transitions
- **Font Analysis**: 
  - Detects all fonts used in the presentation
  - Checks if fonts are installed on your system
  - Shows which slides use each font
  - Reports missing fonts that need to be installed

## Installation

```bash
# Clone the repository
git clone https://github.com/dllmr/PointAssisters.git
```

## Usage

It is recommended to install uv for running these scripts, to avoid the need to manually set up a venv and install required packages.

Command line tool:

```bash
uv run ppta.py presentation.pptx
```

GUI version - file to be chosen via UI:

```bash
uv run qtppta.py
```

## License

[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/)
