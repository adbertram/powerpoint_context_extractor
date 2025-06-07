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
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

from pptx_extractor.utils.common import setup_logging, ensure_directory
from pptx_extractor.notes.extractor import extract_slide_notes
from pptx_extractor.animations.extractor import extract_slide_animations
from pptx_extractor.slides.extractor import extract_slides
from pptx_extractor.recommendations import generate_all_recommendations

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
    parser.add_argument("--recommend", "-r", action="store_true", help="Generate AI-powered usage recommendations for each slide (requires API key)")
    parser.add_argument("--api-key", help="API key for LLM service (can also use ANTHROPIC_API_KEY env var)")
    
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
    
    # Extract animations if requested
    if args.animations or args.all:
        logger.info("Extracting slide animations...")
        animation_data = extract_slide_animations(pptx_path)
    
    # Extract slides if requested
    if args.slides or args.all:
        logger.info("Extracting slides as images...")
        slides_dir = output_path / "slides"
        slides_dir.mkdir(exist_ok=True)
        slide_paths = extract_slides(pptx_path, slides_dir, args.format, args.dpi)
        if slide_paths:
            logger.info(f"Successfully extracted {len(slide_paths)} slides to {slides_dir}")
    
    # Create a unified JSON in the requested format
    logger.info("Creating unified slide content file...")
    slides_data = {"slides": []}
    
    # Determine the maximum number of slides from available data
    max_slides = 0
    if notes_data:
        max_slides = max(max_slides, max([int(key.split('_')[1]) for key in notes_data.keys()]))
    if animation_data:
        max_slides = max(max_slides, max([int(key.split('_')[1]) for key in animation_data.keys()]))
    if slide_paths:
        max_slides = max(max_slides, len(slide_paths))
    
    # Build the slides array
    for slide_num in range(1, max_slides + 1):
        slide_key = f"slide_{slide_num}"
        slide_info = {
            "number": slide_num,
            "title": "",
            "has_animations": False,
            "notes": "",
            "description": ""
        }
        
        # Add notes data if available
        if notes_data and slide_key in notes_data:
            slide_info["title"] = notes_data[slide_key].get("title", "")
            slide_info["notes"] = notes_data[slide_key].get("notes", "")
        
        # Add animation data if available
        if animation_data and slide_key in animation_data:
            slide_info["has_animations"] = animation_data[slide_key].get("has_animations", False)
            # Add animation summary
            if animation_data[slide_key].get("animation_summary"):
                slide_info["animation_summary"] = animation_data[slide_key]["animation_summary"]
            # Add detailed animation information if present
            if animation_data[slide_key].get("animation_details"):
                slide_info["animation_details"] = animation_data[slide_key]["animation_details"]
        
        slides_data["slides"].append(slide_info)
    
    # Generate recommendations if requested
    if args.recommend:
        logger.info("Generating AI-powered usage recommendations...")
        slides_data = generate_all_recommendations(slides_data, args.api_key)
    
    # Save the unified JSON file
    unified_file = save_json_data(slides_data, output_path, "presentation_content.json")
    if unified_file:
        logger.info(f"Successfully saved unified presentation content to {unified_file}")
    
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
    print(f"\nUnified presentation content saved to: {os.path.abspath(args.output)}/presentation_content.json")

if __name__ == "__main__":
    main()
