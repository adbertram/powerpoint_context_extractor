"""
Slide usage recommendations using LLM analysis.
"""

import os
import logging
import json
from typing import Dict, Optional

logger = logging.getLogger(__name__)

# Try to import Anthropic client
try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False
    logger.warning("Anthropic library not installed. Install with: pip install anthropic")

def get_slide_context(slide_data: Dict) -> str:
    """
    Extract relevant context from slide data for LLM analysis.
    
    Args:
        slide_data: Dictionary containing slide information
        
    Returns:
        str: Formatted context string
    """
    context_parts = []
    
    # Add slide number and title
    slide_num = slide_data.get('number', 'Unknown')
    title = slide_data.get('title', '').strip()
    if title:
        context_parts.append(f"Slide {slide_num}: {title}")
    else:
        context_parts.append(f"Slide {slide_num}")
    
    # Add notes if available
    notes = slide_data.get('notes', '').strip()
    if notes:
        context_parts.append(f"\nSpeaker Notes:\n{notes}")
    
    # Add animation summary if available
    anim_summary = slide_data.get('animation_summary', '').strip()
    if anim_summary and anim_summary != "This slide has no animations.":
        context_parts.append(f"\nAnimations:\n{anim_summary}")
    
    # Add brief animation details if present
    anim_details = slide_data.get('animation_details', [])
    if anim_details:
        context_parts.append(f"\nAnimation Effects:")
        for i, anim in enumerate(anim_details[:3]):  # Limit to first 3 for context
            desc = anim.get('description', 'Unknown animation')
            context_parts.append(f"- {desc}")
        if len(anim_details) > 3:
            context_parts.append(f"- ... and {len(anim_details) - 3} more animations")
    
    return "\n".join(context_parts)

def generate_recommendation(slide_data: Dict, api_key: str) -> str:
    """
    Generate usage recommendation for a slide using LLM.
    
    Args:
        slide_data: Dictionary containing slide information
        api_key: API key for Anthropic
        
    Returns:
        str: Usage recommendation paragraph
    """
    if not ANTHROPIC_AVAILABLE:
        return "Recommendation generation unavailable: Anthropic library not installed"
    
    try:
        client = Anthropic(api_key=api_key)
        
        # Extract slide context
        context = get_slide_context(slide_data)
        
        # Create prompt
        prompt = f"""You are a presentation expert analyzing PowerPoint slides. Based on the following slide information, write a single paragraph (3-5 sentences) describing when and how a presenter would want to use this slide. Focus on the practical purpose and ideal usage scenarios.

{context}

Provide a recommendation that:
1. Identifies the slide's primary purpose
2. Suggests when in a presentation it would be most effective
3. Describes what type of content or message it's designed to convey
4. Mentions any special features (animations, layout) that enhance its effectiveness

Write in a professional, helpful tone as if advising a presenter preparing their talk."""

        # Make API call
        response = client.messages.create(
            model="claude-3-haiku-20240307",  # Using Haiku for efficiency
            max_tokens=200,
            temperature=0.7,
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        
        return response.content[0].text.strip()
        
    except Exception as e:
        logger.error(f"Error generating recommendation for slide {slide_data.get('number', 'Unknown')}: {e}")
        return f"Error generating recommendation: {str(e)}"

def generate_all_recommendations(slides_data: Dict, api_key: str) -> Dict:
    """
    Generate recommendations for all slides in the presentation.
    
    Args:
        slides_data: Dictionary containing all slides data
        api_key: API key for Anthropic
        
    Returns:
        Dict: Updated slides data with recommendations
    """
    if not api_key:
        # Check environment variable
        api_key = os.environ.get('ANTHROPIC_API_KEY')
        if not api_key:
            logger.error("No API key provided. Use --api-key or set ANTHROPIC_API_KEY environment variable")
            return slides_data
    
    logger.info(f"Generating recommendations for {len(slides_data.get('slides', []))} slides...")
    
    for slide in slides_data.get('slides', []):
        slide_num = slide.get('number', 'Unknown')
        logger.info(f"Generating recommendation for slide {slide_num}")
        
        recommendation = generate_recommendation(slide, api_key)
        slide['recommended_usage'] = recommendation
    
    return slides_data