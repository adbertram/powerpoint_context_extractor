"""
Slide usage recommendations using LLM analysis.
"""

import os
import logging
import json
from pathlib import Path
from typing import Dict, Optional
from ..config import get_config

logger = logging.getLogger(__name__)

# Try to import Anthropic client
try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False
    logger.warning("Anthropic library not installed. Install with: pip install anthropic")

# Try to import Google Generative AI client
try:
    import google.generativeai as genai
    GOOGLE_AVAILABLE = True
except ImportError:
    GOOGLE_AVAILABLE = False
    logger.warning("Google Generative AI library not installed. Install with: pip install google-generativeai")

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

def load_system_message() -> str:
    """
    Load the system message from the system_message.md file.
    
    Returns:
        str: System message template
    """
    current_dir = Path(__file__).parent
    system_message_path = current_dir / "system_message.md"
    
    try:
        with open(system_message_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # Remove the markdown header and format for use
            lines = content.split('\n')
            # Skip the first header line and empty lines
            filtered_lines = []
            skip_next = False
            for line in lines:
                if line.startswith('# '):
                    continue
                if line.startswith('## '):
                    skip_next = True
                    continue
                if skip_next and line.strip() == '':
                    continue
                if skip_next and line.strip():
                    skip_next = False
                    continue
                filtered_lines.append(line)
            
            return '\n'.join(filtered_lines).strip()
    except Exception as e:
        logger.error(f"Error loading system message: {e}")
        # Fallback to embedded prompt
        return """You are a presentation expert analyzing PowerPoint slides. Based on the following slide information, write a single paragraph (3-5 sentences) describing when and how a presenter would want to use this slide. Focus on the practical purpose and ideal usage scenarios.

{context}

Provide a recommendation that:
1. Identifies the slide's primary purpose
2. Suggests when in a presentation it would be most effective
3. Describes what type of content or message it's designed to convey
4. Mentions any special features (animations, layout) that enhance its effectiveness

Write in a professional, helpful tone as if advising a presenter preparing their talk."""

def create_recommendation_prompt(context: str) -> str:
    """
    Create the recommendation prompt that works for both LLM providers.
    
    Args:
        context: Slide context string
        
    Returns:
        str: Formatted prompt
    """
    system_message = load_system_message()
    return system_message.format(context=context)

def generate_anthropic_recommendation(slide_data: Dict, api_key: str, method: str = "text") -> str:
    """
    Generate usage recommendation using Anthropic's Claude.
    
    Args:
        slide_data: Dictionary containing slide information
        api_key: API key for Anthropic
        method: Recommendation method ("text" or "images")
        
    Returns:
        str: Usage recommendation paragraph
    """
    if not ANTHROPIC_AVAILABLE:
        return "Recommendation generation unavailable: Anthropic library not installed"
    
    try:
        # Get API configuration
        config = get_config()
        api_config = config.get_api_config('anthropic')
        image_settings = config.get('image_settings', {})
        
        client = Anthropic(api_key=api_key)
        
        if method == "images" and "image_path" in slide_data:
            # Use image-based recommendation
            import base64
            
            image_path = slide_data["image_path"]
            if not os.path.exists(image_path):
                return f"Error: Image file not found at {image_path}"
            
            # Read and encode the image
            with open(image_path, "rb") as image_file:
                image_data = base64.b64encode(image_file.read()).decode('utf-8')
            
            # Determine image media type using configuration
            image_ext = os.path.splitext(image_path)[1].lower()
            media_type_map = image_settings.get('supported_media_types', {
                '.png': 'image/png',
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.tiff': 'image/tiff',
                '.bmp': 'image/bmp'
            })
            media_type = media_type_map.get(image_ext, 'image/png')
            
            # Create prompt for image analysis using system message
            system_message = load_system_message()
            image_prompt = system_message.replace("{context}", "Based on this PowerPoint slide image:")
            
            # Make API call with image using configuration
            response = client.messages.create(
                model=api_config.get('model', 'claude-3-haiku-20240307'),
                max_tokens=api_config.get('max_tokens', 200),
                temperature=api_config.get('temperature', 0.7),
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": media_type,
                                    "data": image_data
                                }
                            },
                            {
                                "type": "text",
                                "text": image_prompt
                            }
                        ]
                    }
                ]
            )
        else:
            # Use text-based recommendation (existing functionality)
            context = get_slide_context(slide_data)
            prompt = create_recommendation_prompt(context)

            # Make API call using configuration
            response = client.messages.create(
                model=api_config.get('model', 'claude-3-haiku-20240307'),
                max_tokens=api_config.get('max_tokens', 200),
                temperature=api_config.get('temperature', 0.7),
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
        
        return response.content[0].text.strip()
        
    except Exception as e:
        logger.error(f"Error generating Anthropic recommendation for slide {slide_data.get('number', 'Unknown')}: {e}")
        return f"Error generating recommendation: {str(e)}"

def generate_google_recommendation(slide_data: Dict, api_key: str, method: str = "text") -> str:
    """
    Generate usage recommendation using Google's Gemini.
    
    For available model names, see: https://ai.google.dev/gemini-api/docs/models
    
    Args:
        slide_data: Dictionary containing slide information
        api_key: API key for Google
        method: Recommendation method ("text" or "images")
        
    Returns:
        str: Usage recommendation paragraph
    """
    if not GOOGLE_AVAILABLE:
        return "Recommendation generation unavailable: Google Generative AI library not installed"
    
    try:
        # Get API configuration
        config = get_config()
        api_config = config.get_api_config('google')
        
        # Configure Google AI
        genai.configure(api_key=api_key)
        
        # Create the model using configuration
        model_name = api_config.get('model', 'gemini-2.5-pro-preview-06-05')
        model = genai.GenerativeModel(model_name)
        
        if method == "images" and "image_path" in slide_data:
            # Use image-based recommendation
            from PIL import Image
            
            image_path = slide_data["image_path"]
            if not os.path.exists(image_path):
                return f"Error: Image file not found at {image_path}"
            
            # Load the image
            image = Image.open(image_path)
            
            # Create prompt for image analysis using system message
            system_message = load_system_message()
            image_prompt = system_message.replace("{context}", "Based on this PowerPoint slide image:")
            
            # Generate content with image
            response = model.generate_content([image_prompt, image])
        else:
            # Use text-based recommendation (existing functionality)
            context = get_slide_context(slide_data)
            prompt = create_recommendation_prompt(context)
            
            # Generate content
            response = model.generate_content(prompt)
        
        return response.text.strip()
        
    except Exception as e:
        logger.error(f"Error generating Google recommendation for slide {slide_data.get('number', 'Unknown')}: {e}")
        return f"Error generating recommendation: {str(e)}"

def generate_recommendation(slide_data: Dict, api_key: str, provider: str = "anthropic", method: str = "text") -> str:
    """
    Generate usage recommendation for a slide using the specified LLM provider.
    
    Args:
        slide_data: Dictionary containing slide information
        api_key: API key for the LLM service
        provider: LLM provider to use ("anthropic" or "google")
        method: Recommendation method ("text" or "images")
        
    Returns:
        str: Usage recommendation paragraph
    """
    if provider == "google":
        return generate_google_recommendation(slide_data, api_key, method)
    else:
        return generate_anthropic_recommendation(slide_data, api_key, method)

def generate_all_recommendations(slides_data: Dict, api_key: str, provider: str = "anthropic", method: str = "text") -> Dict:
    """
    Generate recommendations for all slides in the presentation.
    
    Args:
        slides_data: Dictionary containing all slides data
        api_key: API key for the LLM service
        provider: LLM provider to use ("anthropic" or "google")
        method: Recommendation method ("text" or "images")
        
    Returns:
        Dict: Updated slides data with recommendations
    """
    if not api_key:
        # Check environment variable based on provider
        if provider == "google":
            api_key = os.environ.get('GOOGLE_API_KEY')
            if not api_key:
                logger.error("No API key provided. Use --api-key or set GOOGLE_API_KEY environment variable")
                return slides_data
        else:
            api_key = os.environ.get('ANTHROPIC_API_KEY')
            if not api_key:
                logger.error("No API key provided. Use --api-key or set ANTHROPIC_API_KEY environment variable")
                return slides_data
    
    logger.info(f"Generating recommendations using {provider} for {len(slides_data.get('slides', []))} slides...")
    
    for slide in slides_data.get('slides', []):
        slide_num = slide.get('number', 'Unknown')
        logger.info(f"Generating recommendation for slide {slide_num}")
        
        recommendation = generate_recommendation(slide, api_key, provider, method)
        # Only add recommendation if it doesn't start with "Error"
        if not recommendation.startswith("Error"):
            slide['recommended_usage'] = recommendation
    
    return slides_data