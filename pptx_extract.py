#!/usr/bin/env python3
"""
PowerPoint Context Extractor
----------------------------
A comprehensive toolkit for extracting content and metadata from PowerPoint presentations.

This script serves as the main entry point for the PowerPoint Context Extractor toolkit.
It provides a command-line interface for extracting slides, notes, and animations from
PowerPoint presentations.

Usage:
    python pptx_extract.py <pptx_file> [options]

Options:
    --output DIR, -o DIR       Output directory (default: ./output)
    --notes, -n                Extract notes
    --animations, -a           Extract animations
    --slides, -s               Extract slides as images
    --format FORMAT, -f FORMAT Image format for slides (default: png)
    --dpi DPI, -d DPI          Image resolution for slides (default: 300)
    --all                      Extract everything (notes, animations, slides)
    --verbose, -v              Enable verbose logging
    --help, -h                 Show this help message and exit
"""

import os
import sys
import json
import argparse
import logging
from pathlib import Path

from pptx_extractor.utils.common import setup_logging, ensure_directory
from pptx_extractor.notes.extractor import extract_slide_notes
from pptx_extractor.animations.extractor import extract_slide_animations
from pptx_extractor.slides.extractor import extract_slides

def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Extract content and metadata from PowerPoint presentations.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    
    parser.add_argument("pptx_file", help="Path to the PowerPoint file")
    parser.add_argument("--output", "-o", default="./output", help="Output directory (default: ./output)")
    parser.add_argument("--notes", "-n", action="store_true", help="Extract notes")
    parser.add_argument("--animations", "-a", action="store_true", help="Extract animations")
    parser.add_argument("--slides", "-s", action="store_true", help="Extract slides as images")
    parser.add_argument("--format", "-f", default="png", choices=["png", "jpg", "jpeg", "tiff", "bmp"], help="Image format for slides (default: png)")
    parser.add_argument("--dpi", "-d", type=int, default=300, help="Image resolution for slides (default: 300)")
    parser.add_argument("--all", action="store_true", help="Extract everything (notes, animations, slides)")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging")
    
    return parser.parse_args()

def save_json_data(data, output_path, filename):
    """Save data to a JSON file.
    
    Args:
        data: Data to save
        output_path: Output directory
        filename: Output filename
        
    Returns:
        str: Path to the saved file
    """
    file_path = output_path / filename
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
        return str(file_path)
    except Exception as e:
        logging.error(f"Failed to save data to {file_path}: {e}")
        return None

def extract_pptx_content(args):
    """Extract content from a PowerPoint file based on command-line arguments.
    
    Args:
        args: Command-line arguments
        
    Returns:
        tuple: (notes_data, animation_data, slide_paths) or None if extraction failed
    """
    # Set up logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logger = setup_logging(log_level)
    
    # Check if the PowerPoint file exists
    pptx_path = args.pptx_file
    if not os.path.isfile(pptx_path):
        logger.error(f"PowerPoint file not found: {pptx_path}")
        return None
    
    # Create output directory
    output_path = ensure_directory(args.output)
    
    # Initialize result variables
    notes_data = None
    animation_data = None
    slide_paths = None
    
    # Extract notes if requested
    if args.notes or args.all:
        logger.info("Extracting slide notes...")
        notes_data = extract_slide_notes(pptx_path)
        if notes_data:
            notes_file = save_json_data(notes_data, output_path, "slide_notes.json")
            if notes_file:
                logger.info(f"Successfully saved notes information to {notes_file}")
    
    # Extract animations if requested
    if args.animations or args.all:
        logger.info("Extracting slide animations...")
        animation_data = extract_slide_animations(pptx_path)
        if animation_data:
            animation_file = save_json_data(animation_data, output_path, "slide_animations.json")
            if animation_file:
                logger.info(f"Successfully saved animation information to {animation_file}")
    
    # Extract slides if requested
    if args.slides or args.all:
        logger.info("Extracting slides as images...")
        slides_dir = output_path / "slides"
        slides_dir.mkdir(exist_ok=True)
        slide_paths = extract_slides(pptx_path, slides_dir, args.format, args.dpi)
        if slide_paths:
            logger.info(f"Successfully extracted {len(slide_paths)} slides to {slides_dir}")
    
    # Create a combined JSON with notes and animations if both were extracted
    if notes_data and animation_data:
        logger.info("Creating combined content file...")
        combined_data = {}
        for slide_key in notes_data:
            combined_data[slide_key] = {
                'slide_number': notes_data[slide_key]['slide_number'],
                'title': notes_data[slide_key]['title'],
                'notes': notes_data[slide_key]['notes'],
                'animations': animation_data.get(slide_key, {}).get('animations', []),
                'animation_count': animation_data.get(slide_key, {}).get('animation_count', 0)
            }
        
        combined_file = save_json_data(combined_data, output_path, "slide_content.json")
        if combined_file:
            logger.info(f"Successfully saved combined information to {combined_file}")
    
    return notes_data, animation_data, slide_paths

def main():
    """Main entry point."""
    # Parse command-line arguments
    args = parse_arguments()
    
    # If no extraction options are specified, show help and exit
    if not (args.notes or args.animations or args.slides or args.all):
        print("Error: No extraction options specified.")
        print("Please specify at least one of: --notes, --animations, --slides, or --all")
        print("Use --help for more information.")
        sys.exit(1)
    
    # Extract content from PowerPoint file
    result = extract_pptx_content(args)
    
    # Check if extraction was successful
    if result is None:
        sys.exit(1)
    
    # Print summary
    notes_data, animation_data, slide_paths = result
    print("\nExtraction Summary:")
    if notes_data:
        print(f"- Notes extracted from {len(notes_data)} slides")
    if animation_data:
        print(f"- Animations extracted from {len(animation_data)} slides")
    if slide_paths:
        print(f"- {len(slide_paths)} slides extracted as images")
    print(f"\nOutput directory: {os.path.abspath(args.output)}")

if __name__ == "__main__":
    main()
