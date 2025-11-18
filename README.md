# PointAssisters - PowerPoint Analyzer

A suite of tools to analyze PowerPoint presentations (.pptx files) and report on:
- Hidden slides
- Slides featuring animations and transitions
- Embedded media (audio and video files)
- Font usage and missing fonts

## Features

- **Presentation Summary**: Overview of slides, word counts, and presence of effects/media
- **Hidden Slides Detection**: Identifies any slides marked as hidden in the presentation
- **Effects & Media Analysis**:
  - Lists slides containing animations or transitions
  - Detects embedded audio files with filenames
  - Detects embedded video files with filenames
- **Font Analysis**:
  - Detects all fonts used in the presentation (including theme fonts)
  - Checks if fonts are installed on your system
  - Shows which slides use each font
  - Reports missing fonts that need to be installed
  - Flags small fonts below a configurable threshold (GUI only)

## Installation

```bash
# Clone the repository
git clone https://github.com/dllmr/PointAssisters.git
```

## Usage

It is recommended to install [uv](https://docs.astral.sh/uv/) for running these scripts, to avoid the need to manually set up a venv and install required packages. All scripts use PEP 723 inline script metadata for dependency management.

### PowerPoint Analyzer (ppta.py)

**Dual-mode tool**: Runs in CLI mode when a file is provided, or GUI mode when launched without arguments.

**CLI Mode:**
```bash
uv run ppta.py presentation.pptx
```

Displays beautifully formatted analysis results in the terminal using Rich markdown rendering.

**Options:**
- `--threshold N` - Set font size threshold in points (default: 24)
- `--debug` - Enable debug logging

**GUI Mode:**
```bash
uv run ppta.py
```

Opens a Qt-based GUI where you can:
- Select a presentation file via file browser
- Choose which analysis sections to run (Summary, Hidden Slides, Effects & Media, Fonts)
- Configure font size threshold for small font warnings
- View formatted HTML results in the application

### PowerPoint Dump Utility

```bash
uv run pptdump.py presentation.pptx
```

Outputs detailed JSON structure of the entire presentation including all shapes, text runs, theme fonts, media files, and XML elements. Useful for debugging and detailed inspection.

## License

[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/)
