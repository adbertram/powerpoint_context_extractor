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

### Core Dependencies
- Python 3.6+
- python-pptx
- Pillow (for image extraction)
- pdf2image (for slide conversion)
- LibreOffice (for PPTX to PDF conversion)
- Poppler (for PDF to image conversion)

### AI Features (Optional)
- anthropic (for Claude AI recommendations)
- google-generativeai (for Gemini AI recommendations)
- python-dotenv (for environment variable management)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/adbertram/powerpoint_context_extractor.git
   ```

2. Install the package and dependencies:
   ```bash
   # Install basic dependencies
   pip install -r requirements.txt
   
   # Or install the package (includes console script)
   pip install -e .
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
- `--extract TYPE`: What to extract - options: images, notes, animations, all (default: all)
- `--slide-nums RANGES`: Process specific slides (e.g., "1-5,7,10-12")
- `--format FORMAT, -f FORMAT`: Image format for slides (png, jpg, jpeg, tiff, bmp; default: png)
- `--dpi DPI, -d DPI`: Image resolution for slides (default: 300)
- `--recommend, -r`: Generate AI-powered usage recommendations for each slide (requires API key)
- `--recommendation-method METHOD`: Method for recommendations ("text" or "images", default: text)
- `--api-key API_KEY`: API key for LLM service (can also use ANTHROPIC_API_KEY or GOOGLE_API_KEY env var)
- `--llm-provider PROVIDER`: LLM provider to use ("anthropic" or "google", default: anthropic)
- `--config CONFIG`: Path to custom configuration file
- `--verbose, -v`: Enable verbose logging

### Alternative CLI Interface

You can also use the console script (after `pip install -e .`):

```bash
pptx-extract path/to/presentation.pptx [options]
```

Or the basic CLI interface:

```bash
python -m pptx_extractor.cli path/to/presentation.pptx
```

### Examples

#### Extract Everything

```bash
python pptx_extract.py path/to/presentation.pptx --extract all --output ./output_directory
```

#### Extract Only Notes

```bash
python pptx_extract.py path/to/presentation.pptx --extract notes --output ./output_directory
```

#### Extract Only Animations

```bash
python pptx_extract.py path/to/presentation.pptx --extract animations --output ./output_directory
```

#### Extract Slides as Images

```bash
python pptx_extract.py path/to/presentation.pptx --extract images --format png --dpi 300 --output ./output_directory
```

#### Extract Specific Slides

```bash
# Extract slides 1-5, 7, and 10-12
python pptx_extract.py path/to/presentation.pptx --slide-nums "1-5,7,10-12" --output ./output_directory

# Extract only slide 3
python pptx_extract.py path/to/presentation.pptx --slide-nums "3" --output ./output_directory
```

#### Extract with AI-Powered Recommendations

##### Using Anthropic Claude (default)
```bash
# Text-based recommendations (default)
python pptx_extract.py path/to/presentation.pptx --extract all --recommend --api-key YOUR_API_KEY

# Image-based recommendations
python pptx_extract.py path/to/presentation.pptx --extract all --recommend --recommendation-method images --api-key YOUR_API_KEY

# Using environment variable
export ANTHROPIC_API_KEY=YOUR_API_KEY
python pptx_extract.py path/to/presentation.pptx --extract all --recommend
```

##### Using Google Gemini
```bash
# Text-based recommendations
python pptx_extract.py path/to/presentation.pptx --extract all --recommend --llm-provider google --api-key YOUR_API_KEY

# Image-based recommendations
python pptx_extract.py path/to/presentation.pptx --extract all --recommend --llm-provider google --recommendation-method images --api-key YOUR_API_KEY

# Using environment variable
export GOOGLE_API_KEY=YOUR_API_KEY
python pptx_extract.py path/to/presentation.pptx --extract all --recommend --llm-provider google
```

## Output Files

The toolkit generates the following output files:

- **presentation_content.json**: A unified JSON file containing:
  - Slide numbers and titles
  - Notes text
  - Animation details with human-readable descriptions
  - Animation summaries
  - Usage recommendations (when using --recommend option)
- **slides/**: Directory containing extracted slide images (when using image extraction)

## Configuration

The toolkit supports extensive configuration through JSON files. You can customize:

- **API Settings**: Model selection, temperature, token limits for both Anthropic and Google
- **Processing Options**: Timeouts, batch sizes, error handling
- **Output Settings**: File naming conventions, logging levels
- **System Integration**: Dependency management, LibreOffice settings

### Creating a Custom Configuration

```bash
# Use a custom config file
python pptx_extract.py presentation.pptx --config my_config.json
```

The project includes a `config.json.example` file showing all available configuration options. Copy this file to create your own custom configuration:

```bash
# Copy the example config and customize it
cp config.json.example my_config.json
```

Key configuration sections include:

- **cli_defaults**: Override default CLI options (output directory, image format, DPI, etc.)
- **timeouts**: Configure processing timeouts based on slide count and operations
- **api_settings**: Customize AI model settings for both Anthropic and Google providers
- **image_settings**: Control image processing parameters
- **processing**: Fine-tune conversion and processing behavior

See `config.json.example` for the complete configuration structure with all available options and their descriptions.

## Project Structure

```text
powerpoint_context_extractor/
├── pptx_extract.py             # Main entry point script
├── pptx_extractor/             # Package directory
│   ├── __init__.py             # Package initialization
│   ├── cli.py                  # Alternative CLI interface
│   ├── animations/             # Animations extraction module
│   │   ├── __init__.py
│   │   └── extractor.py        # Animations extraction functionality
│   ├── notes/                  # Notes extraction module
│   │   ├── __init__.py
│   │   └── extractor.py        # Notes extraction functionality
│   ├── slides/                 # Slides extraction module
│   │   ├── __init__.py
│   │   └── extractor.py        # Slides extraction functionality
│   ├── recommendations/        # AI-powered recommendations module
│   │   ├── __init__.py
│   │   ├── generator.py        # Recommendation generation logic
│   │   └── system_message.md   # Customizable AI prompts
│   └── utils/                  # Utility functions
│       ├── __init__.py
│       └── common.py           # Common utilities and configuration
├── setup.py                    # Package setup and console scripts
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
