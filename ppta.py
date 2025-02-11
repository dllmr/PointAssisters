# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "python-pptx",
#     "rich",
#     "matplotlib",
# ]
# ///


from pptx import Presentation
from pptx.text.text import _Paragraph
from pptx.shapes.base import BaseShape
from collections import defaultdict
from typing import Dict, List, Set, Tuple
import argparse
from pathlib import Path
from rich.console import Console
from rich.theme import Theme
from rich.table import Table
import matplotlib.font_manager as fm
import logging

INTERNAL_FONT_MARKERS = frozenset({
    '+mj-lt',    # Default Latin font for gothic text
    '+mn-lt',    # Default Latin font for mincho text
    '+body',     # Default body font
    '+major',    # Default major font
    '+minor',    # Default minor font
    '@',         # Font fallback marker
})

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s:%(name)s:%(message)s'
)

logger = logging.getLogger(__name__)

def find_hidden_slides(pptx_path: str) -> List[int]:
    prs = Presentation(pptx_path)
    hidden_slides = []
    
    for slide_num, slide in enumerate(prs.slides, start=1):
        try:
            # Check if slide is marked as hidden
            if hasattr(slide, '_element') and slide._element.get('show') == '0':
                hidden_slides.append(slide_num)
        except Exception as e:
            logger.warning(f"Error checking slide {slide_num}: {e}")
            
    return hidden_slides

def generate_hidden_slides_report(pptx_path: str):
    hidden_slides = find_hidden_slides(pptx_path)
    
    console = Console(theme=Theme({
        "heading": "bold blue"
    }))

    console.print("\n[heading]=== Hidden Slides ===\n")
        
    if hidden_slides:
        console.print("Hidden slides:", (", ".join(str(num) for num in sorted(hidden_slides))))
    else:
        console.print("(no hidden slides found)")

def find_animations_and_transitions(pptx_path: str) -> Tuple[Set[int], Set[int]]:
    """
    Find slides containing transitions or animations in a PowerPoint presentation.
    
    Args:
        pptx_path: Path to the PowerPoint file
        
    Returns:
        Tuple containing:
        - Set of slide numbers with transitions
        - Set of slide numbers with animations
    """
    prs = Presentation(pptx_path)
    slides_with_transitions = set()
    slides_with_animations = set()
    
    for slide_num, slide in enumerate(prs.slides, start=1):
        try:
            # Check for transitions
            transition = slide._element.find('./p:transition', 
                                          {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
            if transition is not None:
                slides_with_transitions.add(slide_num)
            
            # Check for animations
            timing = slide._element.find('./p:timing',
                                      {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
            if timing is not None:
                # Look for any animation elements
                anim_elements = timing.findall('.//p:anim',
                                            {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'})
                anim_elements.extend(timing.findall('.//p:animEffect',
                                                  {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}))
                if anim_elements:
                    slides_with_animations.add(slide_num)
                    
        except Exception as e:
            logger.warning(f"Error processing slide {slide_num}: {e}")
            continue
            
    return slides_with_transitions, slides_with_animations

def print_effects_report(slides_with_transitions: Set[int], slides_with_animations: Set[int]) -> None:
    console = Console(theme=Theme({
        "heading": "bold blue"
    }))

    console.print("\n[heading]=== Transitions and Animations ===\n")
    
    if slides_with_transitions:
        console.print("Slides with transitions:", (", ".join(str(num) for num in sorted(slides_with_transitions))))
    else:
        console.print("(no transitions found)")
        
    if slides_with_animations:
        console.print("Slides with animations:", (", ".join(str(num) for num in sorted(slides_with_animations))))
    else:
        console.print("(no animations found)")

def generate_effects_report(pptx_path: str):
    transitions, animations = find_animations_and_transitions(pptx_path)
    print_effects_report(transitions, animations)
    
def get_system_fonts() -> Set[str]:
    font_list: List[str] = fm.findSystemFonts(fontpaths=None)
    font_names: List[str] = []
    
    for font in font_list:
        try:
            # Attempt to get the font properties
            font_name = fm.FontProperties(fname=font).get_name()
            font_names.append(font_name)
        except Exception as e:
            # Optionally print the error message if you want to debug
            logger.debug(f"Error loading font properties for {font}: {e}")

    return sorted(set(font_names))

def analyze_paragraph_fonts(paragraph: _Paragraph) -> Set[str]:
    """Extract fonts from a paragraph, including runs."""
    fonts = set()

    for run in paragraph.runs:
        try:
            if hasattr(run, 'font') and run.font.name and not is_internal_font(run.font.name):
                fonts.add(run.font.name)

        except Exception as e:
            logger.debug(f"Error analyzing run: {str(e)}")

    return fonts

def analyze_shape_fonts(shape: BaseShape) -> Set[str]:
    """Safely extract fonts from a shape."""
    fonts = set()
    
    try:
        # Handle text frames
        if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                fonts.update(analyze_paragraph_fonts(paragraph))
                
        # Handle tables
        if hasattr(shape, 'has_table') and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        fonts.update(analyze_paragraph_fonts(paragraph))
                        
    except Exception as e:
        logger.debug(f"Error analyzing shape: {str(e)}")
        
    return fonts

def analyze_fonts(pptx_path: str) -> Tuple[Dict[int, Dict[str, Set[str]]], Set[str]]:
    """
    Analyze fonts used in a PowerPoint presentation.
    """
    prs = Presentation(pptx_path)
    font_usage = defaultdict(lambda: defaultdict(set))
    all_fonts = set()
    
    for slide_num, slide in enumerate(prs.slides, start=1):
        try:
            for shape in slide.shapes:
                try:
                    shape_type = f"Text Shape: {shape.name}" if hasattr(shape, 'name') else "Shape"
                    fonts = analyze_shape_fonts(shape)
                    
                    if fonts:
                        font_usage[slide_num][shape_type].update(fonts)
                        all_fonts.update(fonts)
                        
                except Exception as e:
                    logger.debug(f"Error processing shape in slide {slide_num}: {str(e)}")
                    continue
                    
        except Exception as e:
            logger.warning(f"Error processing slide {slide_num}: {str(e)}")
            continue

    return font_usage, all_fonts

def is_internal_font(font_name: str) -> bool:
    if not font_name:
        return True
    font_name = font_name.lower()
    return any(font_name.startswith(marker) for marker in INTERNAL_FONT_MARKERS)

def print_font_report(font_usage: Dict[int, Dict[str, Set[str]]], 
                     all_fonts: Set[str],
                     system_fonts: Set[str]):
    """Print a formatted report showing which slides use each font."""
    # Filter out internal fonts
    all_fonts = {s.strip() for s in all_fonts}
    all_fonts = {f for f in all_fonts if not is_internal_font(f)}
    
    system_fonts = {s.lower().strip() for s in system_fonts}
    
    # Create a mapping of fonts to the slides that use them
    font_to_slides: Dict[str, Set[int]] = {}
    for slide_num, shapes in font_usage.items():
        # Combine all fonts from all shapes in this slide
        slide_fonts = set()
        for shape_fonts in shapes.values():
            slide_fonts.update(shape_fonts)
            
        # Add this slide number to each font's list
        for font in slide_fonts:
            if font and not is_internal_font(font):
                if font not in font_to_slides:
                    font_to_slides[font] = set()
                font_to_slides[font].add(slide_num)

    console = Console(theme=Theme({
        "missing": "red",
        "ok": "green",
        "heading": "bold blue"
    }))
    
    console.print("\n[heading]=== Font Usage ===\n")
    
    if font_to_slides:
        table = Table(show_header=True, header_style="bold")
        table.add_column("Font Name")
        table.add_column("Status")
        table.add_column("Used on Slides")
        
        for font in sorted(font_to_slides.keys()):
            if font:  # Skip None values
                status = "[ok]Installed[/ok]" if font.lower() in system_fonts else "[missing]Missing[/missing]"
                # Convert slide numbers to a readable string
                slides = sorted(font_to_slides[font])
                slides_str = ", ".join(str(slide) for slide in slides)
                table.add_row(font, status, slides_str)
        
        console.print(table)
    else:
        console.print("(no fonts used in presentation)")
    
    # Print summary statistics
    total_fonts = len(font_to_slides)
    missing_fonts = sum(1 for font in font_to_slides if font.lower() not in system_fonts)
    
    console.print("\n[bold]Summary:[/bold]")
    console.print(f"Total unique fonts: {total_fonts}")
    console.print(f"Missing fonts: {missing_fonts}")

def generate_font_report(pptx_path: str):
    # Get system fonts
    system_fonts = get_system_fonts()
    
    # Analyze presentation
    font_usage, all_fonts = analyze_fonts(pptx_path)
    
    # Print report
    print_font_report(font_usage, all_fonts, system_fonts)

def main():
    parser = argparse.ArgumentParser(description='Provide information about a PowerPoint presentation')
    parser.add_argument('pptx_file', type=str, help='Path to the PowerPoint file')
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    args = parser.parse_args()
    
    if args.debug:
        logger.setLevel(logging.DEBUG)
    
    pptx_path = Path(args.pptx_file)
    if not pptx_path.exists():
        logger.error(f"Error: File '{pptx_path}' not found")
        return
    
    try:
        generate_hidden_slides_report(pptx_path)

        generate_effects_report(pptx_path)
            
        generate_font_report(pptx_path)
        
    except Exception as e:
        logger.error(f"Error analyzing presentation: {e}")

if __name__ == "__main__":
    main()
