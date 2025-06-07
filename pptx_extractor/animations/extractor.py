"""
Animations extraction functionality for PowerPoint presentations.
"""

import logging
from pptx import Presentation

from ..utils.common import NAMESPACES, register_namespaces, get_slide_title

logger = logging.getLogger(__name__)

def extract_animation_info(slide_part):
    """
    Extract animation information from a slide's XML part.
    
    Args:
        slide_part: The slide part object from python-pptx
        
    Returns:
        list: List of dictionaries containing animation information
    """
    animations = []
    
    try:
        # Access the slide's XML through the xmlpart property
        if not hasattr(slide_part, 'xmlpart') or not hasattr(slide_part.xmlpart, 'element'):
            logger.debug(f"Slide part does not have xmlpart or element attribute")
            return animations
            
        slide_xml = slide_part.xmlpart.element
        
        # Find timing information
        timing_node = slide_xml.find('.//p:timing', NAMESPACES)
        if timing_node is None:
            return animations
        
        # Find animation sequences
        tn_lt = timing_node.find('.//p:tnLst', NAMESPACES)
        if tn_lt is None:
            return animations
        
        # Process each animation sequence
        for i, par in enumerate(tn_lt.findall('.//p:par', NAMESPACES)):
            ctn = par.find('.//p:cTn', NAMESPACES)
            if ctn is None:
                continue
                
            # Get sequence ID and duration
            seq_id = ctn.get('id', f'unknown_{i}')
            dur = ctn.get('dur', 'unknown')
            
            # Find child animations
            child_tn_lt = ctn.find('.//p:childTnLst', NAMESPACES)
            if child_tn_lt is None:
                continue
                
            # Process each animation effect
            for j, child_par in enumerate(child_tn_lt.findall('.//p:par', NAMESPACES)):
                child_ctn = child_par.find('.//p:cTn', NAMESPACES)
                if child_ctn is None:
                    continue
                    
                # Get effect ID and duration
                effect_id = child_ctn.get('id', f'unknown_effect_{j}')
                effect_dur = child_ctn.get('dur', 'unknown')
                
                # Find target shape
                tgt_el = child_par.find('.//p:tgtEl', NAMESPACES)
                if tgt_el is None:
                    continue
                    
                # Get shape ID
                shape_id_el = tgt_el.find('.//p:spTgt', NAMESPACES)
                shape_id = "unknown"
                if shape_id_el is not None:
                    shape_id = shape_id_el.get('spid', 'unknown')
                
                # Find animation effect
                anim_effect = child_par.find('.//p:animEffect', NAMESPACES)
                effect_type = "unknown"
                if anim_effect is not None:
                    effect_type = anim_effect.get('transition', anim_effect.get('type', 'unknown'))
                
                # Find start conditions
                cond = child_par.find('.//p:cond', NAMESPACES)
                trigger = "unknown"
                delay = "0"
                if cond is not None:
                    trigger = cond.get('evt', 'unknown')
                    delay = cond.get('delay', '0')
                
                # Add animation to list
                animations.append({
                    'sequence_id': seq_id,
                    'effect_id': effect_id,
                    'shape_id': shape_id,
                    'effect_type': effect_type,
                    'trigger': trigger,
                    'delay': delay,
                    'duration': effect_dur
                })
    
    except Exception as e:
        logger.error(f"Error extracting animation info: {e}", exc_info=True)
    
    return animations

def extract_slide_animations(pptx_path):
    """
    Extract animations from all slides in a PowerPoint file.
    
    Args:
        pptx_path (str): Path to the PowerPoint file
        
    Returns:
        dict: Dictionary containing animation information for all slides
    """
    # Register XML namespaces
    register_namespaces()
    
    # Load the presentation
    logger.info(f"Opening PowerPoint file for animation extraction: {pptx_path}")
    try:
        prs = Presentation(pptx_path)
    except Exception as e:
        logger.error(f"Failed to open PowerPoint file: {e}")
        return None
    
    # Dictionary to store animation data
    animation_data = {}
    
    # Process each slide
    for i, slide in enumerate(prs.slides, 1):
        # Get slide title
        title = get_slide_title(slide)
        
        # Extract animations
        animations = extract_animation_info(slide.part)
        
        # Get shape information
        shape_info = {}
        for shape in slide.shapes:
            if shape.shape_id:
                shape_type = "Unknown"
                if hasattr(shape, "shape_type"):
                    shape_type = str(shape.shape_type).replace("MSO_SHAPE_TYPE.", "")
                
                shape_text = ""
                if hasattr(shape, "text") and shape.has_text_frame:
                    shape_text = shape.text.strip()
                
                shape_info[str(shape.shape_id)] = {
                    'type': shape_type,
                    'text': shape_text[:100] + ('...' if len(shape_text) > 100 else '')
                }
        
        # Get slide transition
        transition = "None"
        if hasattr(slide, "slide_layout") and hasattr(slide.slide_layout, "transition"):
            transition = str(slide.slide_layout.transition)
        
        # Add slide information to the dictionary
        animation_data[f"slide_{i}"] = {
            'slide_number': i,
            'title': title,
            'animations': animations,
            'shapes': shape_info,
            'transition': transition,
            'animation_count': len(animations)
        }
        
        logger.info(f"Processed slide {i}: {title[:50]}{'...' if len(title) > 50 else ''} - Animations: {len(animations)}")
    
    return animation_data
