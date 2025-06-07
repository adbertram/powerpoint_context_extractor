#!/usr/bin/env python3
"""
Setup script for PowerPoint Context Extractor.
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="pptx_extractor",
    version="0.1.0",
    author="Adam Bertram",
    author_email="adam@example.com",
    description="A comprehensive toolkit for extracting content and metadata from PowerPoint presentations",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/adbertram/powerpoint_context_extractor",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    install_requires=[
        "python-pptx>=0.6.21",
        "Pillow>=9.5.0",
        "pdf2image>=1.16.3",
    ],
    entry_points={
        "console_scripts": [
            "pptx-extract=pptx_extractor.cli:main",
        ],
    },
)
