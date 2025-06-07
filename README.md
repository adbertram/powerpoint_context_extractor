# PowerPoint Context Extractor

A comprehensive toolkit for extracting content and metadata from PowerPoint presentations, including slide notes, animations, and slide content.

## Overview

PowerPoint Context Extractor is a modular Python toolkit designed to extract and analyze various elements from PowerPoint (.pptx) files. It provides tools to extract:

- **Slide Notes**: Detailed presenter notes from each slide
- **Animations**: Animation sequences, effects, and transitions
- **Slide Content**: Titles, text, and other content elements
- **Slide Images**: Export slides as images in various formats

The toolkit uses a combination of the `python-pptx` library and direct XML parsing to access content that might not be easily accessible through the standard API.

## Features

- **Notes Extraction**: Extract detailed presenter notes from PowerPoint slides
- **Animation Detection**: Identify and document animations, transitions, and effects
- **Content Analysis**: Extract slide titles, content, and structure
- **Combined Output**: Generate comprehensive JSON files containing all extracted information
- **AI-Powered Recommendations**: Generate usage recommendations for each slide using LLM analysis (supports both Anthropic Claude and Google Gemini)
- **Detailed Logging**: Track the extraction process with informative logs
- **Modular Architecture**: Clean separation of concerns for maintainability and extensibility

## Requirements

- Python 3.6+
- python-pptx
- Pillow (for image extraction)
- pdf2image (for slide conversion)
- LibreOffice (for PPTX to PDF conversion)
- Poppler (for PDF to image conversion)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/adbertram/powerpoint_context_extractor.git
   ```

2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Install system dependencies (for slide extraction):
   - LibreOffice: For converting PPTX to PDF
   - Poppler: For converting PDF to images

   On macOS:
   ```bash
   brew install libreoffice poppler
   ```

   On Ubuntu/Debian:
   ```bash
   sudo apt-get install libreoffice poppler-utils
   ```

## Usage

### Main Command-Line Interface

The main script `pptx_extract.py` provides a unified interface for all extraction features:

```bash
python pptx_extract.py path/to/presentation.pptx [options]
```

Options:
- `--output DIR, -o DIR`: Output directory (default: ./output)
- `--notes, -n`: Extract notes
- `--animations, -a`: Extract animations
- `--slides, -s`: Extract slides as images
- `--format FORMAT, -f FORMAT`: Image format for slides (png, jpg, jpeg, tiff, bmp; default: png)
- `--dpi DPI, -d DPI`: Image resolution for slides (default: 300)
- `--all`: Extract everything (notes, animations, slides)
- `--recommend, -r`: Generate AI-powered usage recommendations for each slide (requires API key)
- `--api-key API_KEY`: API key for LLM service (can also use ANTHROPIC_API_KEY or GOOGLE_API_KEY env var)
- `--llm-provider`: LLM provider to use for recommendations ("anthropic" or "google", default: anthropic)
- `--verbose, -v`: Enable verbose logging

### Examples

#### Extract Everything

```bash
python pptx_extract.py path/to/presentation.pptx --all --output ./output_directory
```

#### Extract Only Notes

```bash
python pptx_extract.py path/to/presentation.pptx --notes --output ./output_directory
```

#### Extract Only Animations

```bash
python pptx_extract.py path/to/presentation.pptx --animations --output ./output_directory
```

#### Extract Slides as Images

```bash
python pptx_extract.py path/to/presentation.pptx --slides --format png --dpi 300 --output ./output_directory
```

#### Extract with AI-Powered Recommendations

##### Using Anthropic Claude (default)
```bash
# Using command-line API key
python pptx_extract.py path/to/presentation.pptx --notes --animations --recommend --api-key YOUR_API_KEY --output ./output_directory

# Using environment variable
export ANTHROPIC_API_KEY=YOUR_API_KEY
python pptx_extract.py path/to/presentation.pptx --notes --animations --recommend --output ./output_directory

# Using .env file (create a .env file in the project root)
echo 'ANTHROPIC_API_KEY=YOUR_API_KEY' > .env
python pptx_extract.py path/to/presentation.pptx --notes --animations --recommend --output ./output_directory
```

##### Using Google Gemini
```bash
# Using command-line API key
python pptx_extract.py path/to/presentation.pptx --notes --animations --recommend --llm-provider google --api-key YOUR_API_KEY --output ./output_directory

# Using environment variable
export GOOGLE_API_KEY=YOUR_API_KEY
python pptx_extract.py path/to/presentation.pptx --notes --animations --recommend --llm-provider google --output ./output_directory

# Using .env file (create a .env file in the project root)
echo 'GOOGLE_API_KEY=YOUR_API_KEY' > .env
python pptx_extract.py path/to/presentation.pptx --notes --animations --recommend --llm-provider google --output ./output_directory
```

## Output Files

The toolkit generates the following output files:

- **presentation_content.json**: A unified JSON file containing:
  - Slide numbers and titles
  - Notes text
  - Animation details with human-readable descriptions
  - Animation summaries
  - Usage recommendations (when using --recommend option)
- **slides/**: Directory containing extracted slide images (when using --slides option)


## Project Structure

```text
powerpoint_context_extractor/
├── pptx_extract.py             # Main entry point script
├── pptx_extractor/             # Package directory
│   ├── __init__.py             # Package initialization
│   ├── animations/             # Animations extraction module
│   │   ├── __init__.py
│   │   └── extractor.py        # Animations extraction functionality
│   ├── notes/                  # Notes extraction module
│   │   ├── __init__.py
│   │   └── extractor.py        # Notes extraction functionality
│   ├── slides/                 # Slides extraction module
│   │   ├── __init__.py
│   │   └── extractor.py        # Slides extraction functionality
│   ├── recommendations.py      # AI-powered recommendations module
│   └── utils/                  # Utility functions
│       ├── __init__.py
│       └── common.py           # Common utilities
├── README.md                   # Project documentation
├── LICENSE                     # License file
└── requirements.txt            # Python dependencies
```

## How It Works

The toolkit uses two approaches to extract content:

1. **API-based extraction** using the `python-pptx` library for accessing slide metadata, titles, and basic content
2. **Direct XML parsing** for accessing notes, animations, and other content that might not be easily accessible through the API

For notes extraction, the tool:
1. Opens the PPTX file as a ZIP archive
2. Locates notes XML files in the `ppt/notesSlides/` directory
3. Parses each XML file to find shapes with placeholder type "body"
4. Extracts text from paragraphs and text runs within these shapes
5. Associates the extracted notes with the corresponding slides

## Use Cases

- **Content Analysis**: Analyze the content and structure of PowerPoint presentations
- **Documentation Generation**: Extract presenter notes for documentation purposes
- **Animation Analysis**: Identify slides with complex animations
- **Content Migration**: Extract content for migration to other formats or platforms
- **Accessibility**: Extract notes and content for accessibility purposes

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- [python-pptx](https://python-pptx.readthedocs.io/) library for PowerPoint parsing
- [Pillow](https://pillow.readthedocs.io/) for image processing
- [pdf2image](https://github.com/Belval/pdf2image) for PDF to image conversion
