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
from typing import Dict, List, Set, Tuple, Any, Optional
import argparse
from pathlib import Path
from rich.console import Console
from rich.theme import Theme
from rich.table import Table
import matplotlib.font_manager as fm
import logging
from xml.etree import ElementTree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

INTERNAL_FONT_MARKERS = frozenset({
    '+mj-lt',    # Default Latin font for gothic text
    '+mn-lt',    # Default Latin font for mincho text
    '+body',     # Default body font
    '+major',    # Default major font
    '+minor',    # Default minor font
    '@',         # Font fallback marker
})

THEME_FONT_CODES = {
    '+mj-lt': 'Major Latin',
    '+mn-lt': 'Minor Latin',
    '+mj-ea': 'Major East Asian',
    '+mn-ea': 'Minor East Asian',
    '+mj-cs': 'Major Complex Script',
    '+mn-cs': 'Minor Complex Script',
    '+mj-sym': 'Major Symbol',
    '+mn-sym': 'Minor Symbol',
}

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

def extract_theme_fonts(presentation: Any) -> Dict[str, Any]:
    """Extract theme font information from the presentation."""
    theme_fonts = {}
    
    try:
        # Get the default theme from the first slide master
        if presentation.slide_masters and len(presentation.slide_masters) > 0:
            master = presentation.slide_masters[0]
            
            # Access the theme through the part relationships
            try:
                # Get the master part
                master_part = master.part
                
                # Find theme relationships
                theme_rels = [rel for rel in master_part.rels.values() 
                             if rel.reltype == RT.THEME]
                
                if theme_rels:
                    # Get the first theme part using the relationship ID
                    theme_rel = theme_rels[0]
                    theme_part = master_part.related_part(theme_rel.rId)
                    
                    # Parse the theme XML
                    theme_element = ElementTree.fromstring(theme_part.blob)
                    
                    # Extract font scheme
                    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                    font_scheme_elem = theme_element.find('.//a:fontScheme', ns)
                    
                    if font_scheme_elem is not None:
                        # Get font scheme name
                        scheme_name = font_scheme_elem.get('name', 'Unknown')
                        
                        # Get major font element
                        major_font_elem = font_scheme_elem.find('.//a:majorFont', ns)
                        major_fonts = {}
                        
                        if major_font_elem is not None:
                            latin = major_font_elem.find('.//a:latin', ns)
                            ea = major_font_elem.find('.//a:ea', ns)
                            cs = major_font_elem.find('.//a:cs', ns)
                            sym = major_font_elem.find('.//a:sym', ns)
                            
                            major_fonts = {
                                "latin": latin.get('typeface') if latin is not None else None,
                                "east_asian": ea.get('typeface') if ea is not None else None,
                                "complex_script": cs.get('typeface') if cs is not None else None,
                                "symbol": sym.get('typeface') if sym is not None else None
                            }
                        
                        # Get minor font element
                        minor_font_elem = font_scheme_elem.find('.//a:minorFont', ns)
                        minor_fonts = {}
                        
                        if minor_font_elem is not None:
                            latin = minor_font_elem.find('.//a:latin', ns)
                            ea = minor_font_elem.find('.//a:ea', ns)
                            cs = minor_font_elem.find('.//a:cs', ns)
                            sym = minor_font_elem.find('.//a:sym', ns)
                            
                            minor_fonts = {
                                "latin": latin.get('typeface') if latin is not None else None,
                                "east_asian": ea.get('typeface') if ea is not None else None,
                                "complex_script": cs.get('typeface') if cs is not None else None,
                                "symbol": sym.get('typeface') if sym is not None else None
                            }
                        
                        theme_fonts = {
                            "scheme_name": scheme_name,
                            "major_fonts": major_fonts,
                            "minor_fonts": minor_fonts
                        }
            except Exception as e:
                theme_fonts["error"] = str(e)
    except Exception as e:
        theme_fonts["error"] = str(e)
    
    return theme_fonts

def resolve_theme_font(theme_fonts: Dict[str, Any], theme_code: str) -> Optional[str]:
    """Resolve theme font codes to actual font names."""
    if not theme_code or not theme_code.startswith('+'):
        return None
        
    try:
        if theme_code == "+mj-lt":  # Major font, latin
            return theme_fonts.get("major_fonts", {}).get("latin")
        elif theme_code == "+mn-lt":  # Minor font, latin
            return theme_fonts.get("minor_fonts", {}).get("latin")
        elif theme_code == "+mj-ea":  # Major font, east asian
            return theme_fonts.get("major_fonts", {}).get("east_asian")
        elif theme_code == "+mn-ea":  # Minor font, east asian
            return theme_fonts.get("minor_fonts", {}).get("east_asian")
        elif theme_code == "+mj-cs":  # Major font, complex script
            return theme_fonts.get("major_fonts", {}).get("complex_script")
        elif theme_code == "+mn-cs":  # Minor font, complex script
            return theme_fonts.get("minor_fonts", {}).get("complex_script")
        elif theme_code == "+mj-sym":  # Major font, symbol
            return theme_fonts.get("major_fonts", {}).get("symbol")
        elif theme_code == "+mn-sym":  # Minor font, symbol
            return theme_fonts.get("minor_fonts", {}).get("symbol")
    except Exception:
        pass
        
    return None

def analyze_paragraph_fonts(paragraph: _Paragraph, theme_fonts: Dict[str, Any]) -> Tuple[Set[str], Dict[str, str]]:
    """Extract fonts from a paragraph, including runs and theme fonts."""
    fonts = set()
    theme_font_usage = {}

    for run in paragraph.runs:
        try:
            if hasattr(run, 'font') and run.font.name:
                font_name = run.font.name
                
                # Check if it's a theme font
                if font_name.startswith('+'):
                    theme_type = THEME_FONT_CODES.get(font_name, font_name)
                    resolved_font = resolve_theme_font(theme_fonts, font_name)
                    
                    if resolved_font:
                        theme_font_usage[theme_type] = resolved_font
                        fonts.add(resolved_font)
                    else:
                        theme_font_usage[theme_type] = "MISSING"
                elif not is_internal_font(font_name):
                    fonts.add(font_name)

        except Exception as e:
            logger.debug(f"Error analyzing run: {str(e)}")

    return fonts, theme_font_usage

def analyze_shape_fonts(shape: BaseShape, theme_fonts: Dict[str, Any]) -> Tuple[Set[str], Dict[str, str]]:
    """Safely extract fonts from a shape, including theme fonts."""
    fonts = set()
    theme_font_usage = {}
    
    try:
        # Handle text frames
        if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                para_fonts, para_theme_fonts = analyze_paragraph_fonts(paragraph, theme_fonts)
                fonts.update(para_fonts)
                theme_font_usage.update(para_theme_fonts)
                
        # Handle tables
        if hasattr(shape, 'has_table') and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        para_fonts, para_theme_fonts = analyze_paragraph_fonts(paragraph, theme_fonts)
                        fonts.update(para_fonts)
                        theme_font_usage.update(para_theme_fonts)
                        
    except Exception as e:
        logger.debug(f"Error analyzing shape: {str(e)}")
        
    return fonts, theme_font_usage

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
                    fonts = set()
                    
                    # Handle text frames
                    if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                try:
                                    if hasattr(run, 'font') and run.font.name:
                                        font_name = run.font.name
                                        if not is_internal_font(font_name):
                                            fonts.add(font_name)
                                except Exception as e:
                                    logger.debug(f"Error analyzing run: {str(e)}")
                    
                    # Handle tables
                    if hasattr(shape, 'has_table') and shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        try:
                                            if hasattr(run, 'font') and run.font.name:
                                                font_name = run.font.name
                                                if not is_internal_font(font_name):
                                                    fonts.add(font_name)
                                        except Exception as e:
                                            logger.debug(f"Error analyzing run: {str(e)}")
                    
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
    return any(font_name.startswith(marker.lower()) for marker in INTERNAL_FONT_MARKERS)

def print_font_report(font_usage: Dict[int, Dict[str, Set[str]]], 
                     all_fonts: Set[str],
                     system_fonts: Set[str],
                     presentation: Any):
    """Print a formatted report showing font usage and theme fonts."""
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
        "heading": "bold blue",
        "theme": "bold magenta"
    }))
    
    # Print regular font usage
    console.print("\n[heading]=== Regular Font Usage ===\n")
    
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
        console.print("(no regular fonts used in presentation)")
    
    # Print theme fonts
    console.print("\n[heading]=== Theme Fonts ===\n")
    
    # Extract theme fonts
    theme_fonts = extract_theme_fonts(presentation)
    
    if "error" in theme_fonts:
        console.print(f"[missing]Error accessing theme fonts: {theme_fonts['error']}[/missing]")
    else:
        theme_table = Table(show_header=True, header_style="bold")
        theme_table.add_column("Font Type")
        theme_table.add_column("Font Name")
        theme_table.add_column("Status")
        
        # Process major fonts
        major_fonts = theme_fonts.get("major_fonts", {})
        for script, font in major_fonts.items():
            if font:
                status = "[ok]Installed[/ok]" if font.lower() in system_fonts else "[missing]Missing[/missing]"
                theme_table.add_row(
                    f"[theme]Major {script.replace('_', ' ').title()}[/theme]",
                    font,
                    status
                )
        
        # Process minor fonts
        minor_fonts = theme_fonts.get("minor_fonts", {})
        for script, font in minor_fonts.items():
            if font:
                status = "[ok]Installed[/ok]" if font.lower() in system_fonts else "[missing]Missing[/missing]"
                theme_table.add_row(
                    f"[theme]Minor {script.replace('_', ' ').title()}[/theme]",
                    font,
                    status
                )
        
        if major_fonts or minor_fonts:
            console.print(f"Theme scheme name: {theme_fonts.get('scheme_name', 'Unknown')}")
            console.print(theme_table)
        else:
            console.print("(no theme fonts defined)")
    
    # Print summary statistics
    total_fonts = len(font_to_slides)
    missing_fonts = sum(1 for font in font_to_slides if font.lower() not in system_fonts)
    
    total_theme_fonts = sum(
        len([f for f in fonts.values() if f]) 
        for fonts in [theme_fonts.get("major_fonts", {}), theme_fonts.get("minor_fonts", {})]
    )
    missing_theme_fonts = sum(
        len([f for f in fonts.values() if f and f.lower() not in system_fonts])
        for fonts in [theme_fonts.get("major_fonts", {}), theme_fonts.get("minor_fonts", {})]
    )
    
    console.print("\n[bold]Summary:[/bold]")
    console.print(f"Total regular fonts: {total_fonts}")
    console.print(f"Missing regular fonts: {missing_fonts}")
    console.print(f"Total theme fonts: {total_theme_fonts}")
    console.print(f"Missing theme fonts: {missing_theme_fonts}")

def generate_font_report(pptx_path: str):
    # Get system fonts
    system_fonts = get_system_fonts()
    
    # Open presentation
    prs = Presentation(pptx_path)
    
    # Analyze presentation
    font_usage, all_fonts = analyze_fonts(pptx_path)
    
    # Print report
    print_font_report(font_usage, all_fonts, system_fonts, prs)

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
