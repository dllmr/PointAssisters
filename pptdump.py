# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "python-pptx",
# ]
# ///


import json
import sys
from pathlib import Path
from pptx import Presentation
from typing import Any, Dict, List
from xml.etree import ElementTree


def resolve_theme_font(shape: Any, theme_code: str) -> str:
    """Resolve theme font codes to actual font names."""
    try:
        if not theme_code:
            return None

        # Get the theme from the slide master
        theme = shape.part.slide.slide_layout.slide_master.theme

        # Handle theme codes
        if theme_code == "+mn-lt":  # Minor font, latin
            return theme.font_scheme.minor_font.latin
        elif theme_code == "+mj-lt":  # Major font, latin
            return theme.font_scheme.major_font.latin
        elif theme_code == "+mn-ea":  # Minor font, east asian
            return theme.font_scheme.minor_font.east_asian
        elif theme_code == "+mj-ea":  # Major font, east asian
            return theme.font_scheme.major_font.east_asian
        elif theme_code == "+mn-cs":  # Minor font, complex script
            return theme.font_scheme.minor_font.complex_script
        elif theme_code == "+mj-cs":  # Major font, complex script
            return theme.font_scheme.major_font.complex_script
        else:
            return theme_code  # If it's not a theme code, return as is
    except Exception:
        return theme_code  # Return original code if resolution fails

def shape_to_dict(shape: Any) -> Dict[str, Any]:
    """Convert a shape object to a dictionary of its properties."""
    shape_dict = {
        "name": shape.name,
        "shape_type": str(shape.shape_type),
        "width": shape.width,
        "height": shape.height,
        "left": shape.left,
        "top": shape.top,
    }

    # Handle text if present
    if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
        shape_dict["text_frame"] = {
            "text": shape.text,
            "default_font": None,
            "theme_font": None,
            "resolved_font": None,
            "paragraphs": []
        }

        # Try to get shape-level font defaults from XML
        if hasattr(shape, '_element'):
            ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            shape_props = shape._element.find('.//a:pPr', ns)
            if shape_props is not None:
                latin_font = shape_props.find('.//a:latin', ns)
                if latin_font is not None:
                    theme_font = latin_font.get('typeface')
                    shape_dict["text_frame"]["theme_font"] = theme_font
                    shape_dict["text_frame"]["resolved_font"] = resolve_theme_font(shape, theme_font)

        # Try to get text frame level defaults
        try:
            if hasattr(shape.text_frame, 'properties'):
                shape_dict["text_frame"]["properties"] = {
                    "margin_left": shape.text_frame.margin_left,
                    "margin_right": shape.text_frame.margin_right,
                    "margin_top": shape.text_frame.margin_top,
                    "margin_bottom": shape.text_frame.margin_bottom,
                    "vertical_anchor": str(shape.text_frame.vertical_anchor),
                    "word_wrap": shape.text_frame.word_wrap,
                    "auto_size": str(shape.text_frame.auto_size) if hasattr(shape.text_frame, 'auto_size') else None
                }
        except Exception as e:
            shape_dict["text_frame"]["properties_error"] = str(e)

        for p in shape.text_frame.paragraphs:
            para_dict = {
                "text": p.text,
                "level": p.level,
                "alignment": str(p.alignment) if hasattr(p, 'alignment') else None,
                "runs": []
            }

            # Try to get paragraph level font defaults
            try:
                if hasattr(p, 'font'):
                    para_dict["default_font"] = {
                        "name": p.font.name,
                        "size": p.font.size.pt if p.font.size else None,
                        "bold": p.font.bold,
                        "italic": p.font.italic,
                        "underline": p.font.underline
                    }
                # Try to get font info from XML
                if hasattr(p, '_element'):
                    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                    para_props = p._element.find('.//a:pPr', ns)
                    if para_props is not None:
                        latin_font = para_props.find('.//a:latin', ns)
                        if latin_font is not None:
                            theme_font = latin_font.get('typeface')
                            para_dict["theme_font"] = theme_font
                            para_dict["resolved_font"] = resolve_theme_font(shape, theme_font)
            except Exception as e:
                para_dict["font_error"] = str(e)

            for run in p.runs:
                run_dict = {
                    "text": run.text,
                    "font": None,
                    "theme_font": None,
                    "resolved_font": None
                }
                try:
                    if hasattr(run, 'font'):
                        run_dict["font"] = {
                            "name": run.font.name,
                            "size": run.font.size.pt if run.font.size else None,
                            "bold": run.font.bold,
                            "italic": run.font.italic,
                            "underline": run.font.underline,
                        }
                    # Try to get run-level font info from XML
                    if hasattr(run, '_element'):
                        ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                        run_props = run._element.find('.//a:rPr', ns)
                        if run_props is not None:
                            latin_font = run_props.find('.//a:latin', ns)
                            if latin_font is not None:
                                theme_font = latin_font.get('typeface')
                                run_dict["theme_font"] = theme_font
                                run_dict["resolved_font"] = resolve_theme_font(shape, theme_font)
                except Exception as e:
                    run_dict["font_error"] = str(e)
                para_dict["runs"].append(run_dict)

            shape_dict["text_frame"]["paragraphs"].append(para_dict)

    # Handle table if present
    if hasattr(shape, 'has_table') and shape.has_table:
        shape_dict["table"] = {
            "rows": len(shape.table.rows),
            "columns": len(shape.table.columns),
            "cells": [
                [
                    {
                        "text": cell.text,
                        "location": f"row {row_idx}, col {col_idx}"
                    }
                    for col_idx, cell in enumerate(row.cells)
                ]
                for row_idx, row in enumerate(shape.table.rows)
            ]
        }

    # Handle image if present
    if hasattr(shape, 'image'):
        try:
            shape_dict["image"] = {
                "filename": shape.image.filename,
                "content_type": shape.image.content_type,
                "size": shape.image.size,
            }
        except AttributeError:
            shape_dict["image"] = None

    # Add placeholder information if available
    if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
        try:
            shape_dict["placeholder"] = {
                "type": str(shape.placeholder_format.type),
                "idx": shape.placeholder_format.idx
            }
        except AttributeError:
            pass

    return shape_dict

def slide_to_dict(slide: Any, slide_index: int) -> Dict[str, Any]:
    """Convert a slide object to a dictionary of its properties."""
    slide_dict = {
        "slide_number": slide_index + 1,
        "shapes": [],
        "slide_id": slide.slide_id,
    }

    # Check if slide is hidden
    if hasattr(slide, '_element'):
        slide_dict["hidden"] = slide._element.get('show') == '0'

    # Get slide layout info
    if hasattr(slide, 'slide_layout'):
        slide_dict["layout"] = {
            "name": slide.slide_layout.name,
        }

    # Get background info if available
    try:
        background = slide._element.find('.//p:bg',
            {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
        if background is not None:
            slide_dict["background"] = ElementTree.tostring(background).decode()
    except Exception:
        slide_dict["background"] = None

    # Check for transitions
    try:
        transition = slide._element.find('.//p:transition',
            {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
        if transition is not None:
            slide_dict["has_transition"] = True
            slide_dict["transition_xml"] = ElementTree.tostring(transition).decode()
    except Exception:
        slide_dict["has_transition"] = False

    # Check for animations
    try:
        timing = slide._element.find('.//p:timing',
            {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
        if timing is not None:
            slide_dict["has_animations"] = True
            slide_dict["timing_xml"] = ElementTree.tostring(timing).decode()
    except Exception:
        slide_dict["has_animations"] = False

    # Convert each shape
    for shape in slide.shapes:
        try:
            shape_dict = shape_to_dict(shape)
            slide_dict["shapes"].append(shape_dict)
        except Exception as e:
            slide_dict["shapes"].append({
                "error": f"Failed to convert shape: {str(e)}",
                "shape_name": getattr(shape, 'name', 'unknown')
            })

    return slide_dict

def presentation_to_dict(pptx_path: Path) -> Dict[str, Any]:
    """Convert entire presentation to a dictionary."""
    prs = Presentation(pptx_path)

    # Basic presentation properties
    pres_dict = {
        "metadata": {
            "slides_count": len(prs.slides),
            "core_properties": {
                "author": prs.core_properties.author,
                "created": str(prs.core_properties.created) if prs.core_properties.created else None,
                "modified": str(prs.core_properties.modified) if prs.core_properties.modified else None,
                "title": prs.core_properties.title,
                "subject": prs.core_properties.subject,
                "keywords": prs.core_properties.keywords,
                "comments": prs.core_properties.comments,
                "category": prs.core_properties.category,
            }
        },
        "slides": []
    }

    # Convert each slide
    for idx, slide in enumerate(prs.slides):
        try:
            slide_dict = slide_to_dict(slide, idx)
            pres_dict["slides"].append(slide_dict)
        except Exception as e:
            pres_dict["slides"].append({
                "error": f"Failed to convert slide {idx + 1}: {str(e)}",
                "slide_number": idx + 1
            })

    return pres_dict

def main():
    if len(sys.argv) != 2:
        print("Usage: python script.py <path_to_pptx>")
        sys.exit(1)

    pptx_path = Path(sys.argv[1])
    if not pptx_path.exists():
        print(f"Error: File '{pptx_path}' does not exist")
        sys.exit(1)

    if pptx_path.suffix.lower() != '.pptx':
        print("Error: File must be a .pptx file")
        sys.exit(1)

    try:
        pres_dict = presentation_to_dict(pptx_path)
        print(json.dumps(pres_dict, indent=2, default=str))
    except Exception as e:
        print(f"Error processing presentation: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
