"""
Common utilities for PowerPoint extraction.
"""

import os
import logging
import xml.etree.ElementTree as ET
from pathlib import Path

# Define XML namespaces used in PPTX files
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p14': 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
}

def setup_logging(level=logging.INFO):
    """Set up logging configuration.
    
    Args:
        level: Logging level (default: INFO)
        
    Returns:
        Logger object
    """
    logging.basicConfig(level=level, format='%(asctime)s - %(levelname)s - %(message)s')
    return logging.getLogger(__name__)

def register_namespaces():
    """Register XML namespaces for parsing."""
    for prefix, uri in NAMESPACES.items():
        ET.register_namespace(prefix, uri)

def ensure_directory(directory_path):
    """Ensure that a directory exists, creating it if necessary.
    
    Args:
        directory_path: Path to the directory
        
    Returns:
        Path object for the directory
    """
    path = Path(directory_path)
    path.mkdir(parents=True, exist_ok=True)
    return path

def sanitize_filename(filename):
    """Sanitize a filename by removing invalid characters.
    
    Args:
        filename: Original filename
        
    Returns:
        Sanitized filename
    """
    # Replace invalid characters with underscores
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Limit length and trim whitespace
    return filename.strip()[:100]

def get_slide_title(slide):
    """Extract the title from a slide.
    
    Args:
        slide: Slide object from python-pptx
        
    Returns:
        Slide title or "Untitled" if no title is found
    """
    title = "Untitled"
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.has_text_frame:
            if shape.text.strip():
                title = shape.text.strip().replace('\n', ' ')
                break
    return title
