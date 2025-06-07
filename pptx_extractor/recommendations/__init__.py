"""
Recommendations module for generating AI-powered slide usage recommendations.
"""

from .generator import (
    generate_recommendation,
    generate_all_recommendations,
    get_slide_context
)

__all__ = [
    'generate_recommendation',
    'generate_all_recommendations',
    'get_slide_context'
]