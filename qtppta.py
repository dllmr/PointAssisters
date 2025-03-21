# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "markdown",
#     "matplotlib",
#     "python-pptx",
#     "PySide6",
# ]
# ///


import logging
import markdown
import sys
from collections import defaultdict
from pathlib import Path
from pptx import Presentation
from pptx.shapes.base import BaseShape
from pptx.text.text import _Paragraph
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QFileDialog, QLineEdit, 
                            QTextEdit, QCheckBox, QGroupBox, QStatusBar)
from typing import Dict, List, Set, Tuple, Any
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

def generate_hidden_slides_report(pptx_path: str) -> str:
    hidden_slides = find_hidden_slides(pptx_path)
    
    result = ""

    result += "## Hidden Slides\n"
        
    if hidden_slides:
        result += "Hidden slides: " + (", ".join(str(num) for num in sorted(hidden_slides))) + "\n"
    else:
        result += "(no hidden slides found)\n"
    result += "***\n"

    return result

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

def format_effects_report(slides_with_transitions: Set[int], slides_with_animations: Set[int]) -> str:
    result = ""

    result += "## Transitions and Animations\n"
    
    if slides_with_transitions:
        result += "Slides with transitions: " + (", ".join(str(num) for num in sorted(slides_with_transitions))) + "<br />\n"
    else:
        result += "(no transitions found)<br />\n"
        
    if slides_with_animations:
        result += "Slides with animations: " + (", ".join(str(num) for num in sorted(slides_with_animations))) + "\n"
    else:
        result += "(no animations found)\n"
    result += "***\n"

    return result

def generate_effects_report(pptx_path: str) -> str:
    transitions, animations = find_animations_and_transitions(pptx_path)
    return format_effects_report(transitions, animations)
    
def get_system_fonts() -> Set[str]:
    import matplotlib.font_manager as fm
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

def analyze_paragraph_fonts(paragraph: _Paragraph) -> Dict[str, bool]:
    """
    Extract fonts from a paragraph, including runs.
    
    Returns:
        Dict mapping font names to whether they contain visible text (True) or only whitespace (False)
    """
    fonts = {}

    for run in paragraph.runs:
        try:
            if hasattr(run, 'font') and run.font.name and not is_internal_font(run.font.name):
                font_name = run.font.name
                # Check if this run contains non-whitespace characters
                has_visible_text = bool(run.text.strip())
                
                # If font already exists in our map and previously had visible text, keep it that way
                if font_name in fonts:
                    fonts[font_name] = fonts[font_name] or has_visible_text
                else:
                    fonts[font_name] = has_visible_text

        except Exception as e:
            logger.debug(f"Error analyzing run: {str(e)}")

    return fonts

def analyze_shape_fonts(shape: BaseShape) -> Dict[str, bool]:
    """
    Safely extract fonts from a shape, tracking whether each font has visible text.
    
    Returns:
        Dict mapping font names to whether they contain visible text (True) or only whitespace (False)
    """
    fonts = {}
    
    try:
        # Handle text frames
        if hasattr(shape, 'has_text_frame') and shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                paragraph_fonts = analyze_paragraph_fonts(paragraph)
                # Merge results, keeping track of visible text status
                for font_name, has_visible_text in paragraph_fonts.items():
                    if font_name in fonts:
                        fonts[font_name] = fonts[font_name] or has_visible_text
                    else:
                        fonts[font_name] = has_visible_text
                
        # Handle tables
        if hasattr(shape, 'has_table') and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph_fonts = analyze_paragraph_fonts(paragraph)
                        # Merge results, keeping track of visible text status
                        for font_name, has_visible_text in paragraph_fonts.items():
                            if font_name in fonts:
                                fonts[font_name] = fonts[font_name] or has_visible_text
                            else:
                                fonts[font_name] = has_visible_text
                        
    except Exception as e:
        logger.debug(f"Error analyzing shape: {str(e)}")
        
    return fonts

def analyze_fonts(pptx_path: str) -> Tuple[Dict[int, Dict[str, Dict[str, bool]]], Dict[str, bool]]:
    """
    Analyze fonts used in a PowerPoint presentation.
    
    Returns:
        Tuple containing:
        - Dictionary mapping slide numbers to shape types to font usage info
        - Dictionary mapping all fonts to whether they have visible text in the presentation
    """
    prs = Presentation(pptx_path)
    font_usage = defaultdict(lambda: defaultdict(dict))
    all_fonts_visibility = {}
    
    for slide_num, slide in enumerate(prs.slides, start=1):
        try:
            for shape in slide.shapes:
                try:
                    shape_type = f"Text Shape: {shape.name}" if hasattr(shape, 'name') else "Shape"
                    fonts = analyze_shape_fonts(shape)
                    
                    if fonts:
                        font_usage[slide_num][shape_type] = fonts
                        
                        # Update global font visibility tracking
                        for font_name, has_visible_text in fonts.items():
                            if font_name in all_fonts_visibility:
                                all_fonts_visibility[font_name] = all_fonts_visibility[font_name] or has_visible_text
                            else:
                                all_fonts_visibility[font_name] = has_visible_text
                        
                except Exception as e:
                    logger.debug(f"Error processing shape in slide {slide_num}: {str(e)}")
                    continue
                    
        except Exception as e:
            logger.warning(f"Error processing slide {slide_num}: {str(e)}")
            continue

    return font_usage, all_fonts_visibility

def is_internal_font(font_name: str) -> bool:
    if not font_name:
        return True
    font_name = font_name.lower()
    return any(font_name.startswith(marker) for marker in INTERNAL_FONT_MARKERS)

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

def format_font_report(font_usage: Dict[int, Dict[str, Dict[str, bool]]], 
                     all_fonts_visibility: Dict[str, bool],
                     system_fonts: Set[str],
                     presentation: Any) -> str:
    """Create a formatted report showing font usage and theme fonts."""
    result = ""

    # Add CSS styling for tables
    result += """<style>
table {
    border-collapse: collapse;
    margin: 10px 0;
    width: 100%;
}
th, td {
    padding: 8px 16px;
    text-align: left;
    border: 1px solid #666;
}
th {
    font-weight: bold;
}
.whitespace-only {
    color: #888;
    font-style: italic;
}
</style>
"""

    # Filter out internal fonts
    all_fonts_visibility = {f.strip(): v for f, v in all_fonts_visibility.items() if not is_internal_font(f)}
    
    system_fonts = {s.lower().strip() for s in system_fonts}
    
    # Create a mapping of fonts to the slides that use them, with visibility information
    font_to_slides: Dict[str, Dict[int, bool]] = {}
    for slide_num, shapes in font_usage.items():
        # Process all fonts from all shapes in this slide
        for shape_type, fonts_info in shapes.items():
            for font, has_visible_text in fonts_info.items():
                if font and not is_internal_font(font):
                    if font not in font_to_slides:
                        font_to_slides[font] = {}
                    # Store whether this font has visible text on this slide
                    if slide_num in font_to_slides[font]:
                        font_to_slides[font][slide_num] = font_to_slides[font][slide_num] or has_visible_text
                    else:
                        font_to_slides[font][slide_num] = has_visible_text

    result += "## Regular Font Usage\n"
    
    if font_to_slides:
        result += "<table>\n"
        result += "<tr><th>Font Name</th><th>Status</th><th>Used on Slides</th><th>Notes</th></tr>\n"
        
        for font in sorted(font_to_slides.keys()):
            if font:  # Skip None values
                status = "✅ Installed" if font.lower() in system_fonts else "❌ Missing"
                
                # Convert slide numbers to a readable string, marking whitespace-only slides
                slides_info = sorted(font_to_slides[font].items())
                slide_parts = []
                
                for slide_num, has_visible_text in slides_info:
                    if has_visible_text:
                        slide_parts.append(str(slide_num))
                    else:
                        slide_parts.append(f"<span class='whitespace-only'>{slide_num}*</span>")
                
                slides_str = ", ".join(slide_parts)
                
                # Determine if this font is only used for whitespace across all slides
                whitespace_only = not any(has_visible_text for _, has_visible_text in slides_info)
                
                # Add note about whitespace usage
                notes = ""
                if whitespace_only:
                    notes = "<span class='whitespace-only'>Used only for whitespace</span>"
                elif any(not has_visible_text for _, has_visible_text in slides_info):
                    notes = "<span class='whitespace-only'>* = whitespace only on marked slides</span>"
                
                result += f"<tr><td>{font}</td><td>{status}</td><td>{slides_str}</td><td>{notes}</td></tr>\n"
        
        result += "</table>\n"
    else:
        result += "(no regular fonts used in presentation)\n"
    
    # Print theme fonts
    result += "\n## Theme Fonts\n"
    
    # Extract theme fonts
    theme_fonts = extract_theme_fonts(presentation)
    
    if "error" in theme_fonts:
        result += f"Error accessing theme fonts: {theme_fonts['error']}\n"
    else:
        if theme_fonts.get("scheme_name"):
            result += f"Theme scheme name: {theme_fonts['scheme_name']}\n\n"
        
        result += "<table>\n"
        result += "<tr><th>Font Type</th><th>Font Name</th><th>Status</th></tr>\n"
        
        # Process major fonts
        major_fonts = theme_fonts.get("major_fonts", {})
        for script, font in major_fonts.items():
            if font:
                status = "✅ Installed" if font.lower() in system_fonts else "❌ Missing"
                result += f"<tr><td>Major {script.replace('_', ' ').title()}</td><td>{font}</td><td>{status}</td></tr>\n"
        
        # Process minor fonts
        minor_fonts = theme_fonts.get("minor_fonts", {})
        for script, font in minor_fonts.items():
            if font:
                status = "✅ Installed" if font.lower() in system_fonts else "❌ Missing"
                result += f"<tr><td>Minor {script.replace('_', ' ').title()}</td><td>{font}</td><td>{status}</td></tr>\n"
        
        result += "</table>\n"
        
        if not (major_fonts or minor_fonts):
            result += "(no theme fonts defined)\n"
    
    # Print summary statistics
    total_fonts = len(font_to_slides)
    missing_fonts = sum(1 for font in font_to_slides if font.lower() not in system_fonts)
    
    # Count whitespace-only fonts
    whitespace_only_fonts = sum(
        1 for font in font_to_slides 
        if not any(has_visible_text for _, has_visible_text in font_to_slides[font].items())
    )
    
    total_theme_fonts = sum(
        len([f for f in fonts.values() if f]) 
        for fonts in [theme_fonts.get("major_fonts", {}), theme_fonts.get("minor_fonts", {})]
    )
    missing_theme_fonts = sum(
        len([f for f in fonts.values() if f and f.lower() not in system_fonts])
        for fonts in [theme_fonts.get("major_fonts", {}), theme_fonts.get("minor_fonts", {})]
    )
    
    result += "\n## Summary\n"
    result += f"Total regular fonts: {total_fonts}<br />\n"
    result += f"Missing regular fonts: {missing_fonts}<br />\n"
    result += f"Fonts used only for whitespace: {whitespace_only_fonts}<br />\n"
    result += f"Total theme fonts: {total_theme_fonts}<br />\n"
    result += f"Missing theme fonts: {missing_theme_fonts}\n"
    result += "***\n"

    return result

def generate_font_report(pptx_path: str) -> str:
    # Get system fonts
    system_fonts = get_system_fonts()
    
    # Open presentation
    prs = Presentation(pptx_path)
    
    # Analyze presentation
    font_usage, all_fonts_visibility = analyze_fonts(pptx_path)
    
    # Format report
    return format_font_report(font_usage, all_fonts_visibility, system_fonts, prs)

class PowerPointAnalyzerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PointAssisters - PowerPoint Analyzer")
        self.resize(800, 600)

        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # File selection
        file_layout = QHBoxLayout()
        self.file_entry = QLineEdit()
        self.file_entry.setPlaceholderText("Select a PowerPoint file...")
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_entry)
        file_layout.addWidget(browse_button)
        layout.addLayout(file_layout)

        # Analysis options
        options_group = QGroupBox("Analysis Options")
        options_layout = QHBoxLayout()
        self.hidden_check = QCheckBox("Hidden Slides")
        self.effects_check = QCheckBox("Effects")
        self.fonts_check = QCheckBox("Fonts")
        self.hidden_check.setChecked(True)
        self.effects_check.setChecked(True)
        self.fonts_check.setChecked(True)
        options_layout.addWidget(self.hidden_check)
        options_layout.addWidget(self.effects_check)
        options_layout.addWidget(self.fonts_check)
        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        # Analyze button
        analyze_button = QPushButton("Analyze")
        analyze_button.clicked.connect(self.analyze)
        layout.addWidget(analyze_button)

        # Results area
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        layout.addWidget(self.results_text)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

    def browse_file(self):
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select PowerPoint File",
            "",
            "PowerPoint files (*.pptx);;All files (*.*)"
        )
        if filename:
            self.file_entry.setText(filename)
            self.results_text.clear()
            self.status_bar.clearMessage()

    def analyze(self):
        file_path = self.file_entry.text()
        if not file_path:
            self.status_bar.showMessage("Please select a file first")
            return

        if not Path(file_path).exists():
            self.status_bar.showMessage("Selected file does not exist")
            return

        self.results_text.clear()
        self.status_bar.showMessage("Analyzing...")
        QApplication.processEvents()  # Ensure UI updates

        try:
            # Capture output
            output = ""
            if self.hidden_check.isChecked():
                output += generate_hidden_slides_report(file_path)
            if self.effects_check.isChecked():
                output += generate_effects_report(file_path)
            if self.fonts_check.isChecked():
                output += generate_font_report(file_path)

            # Display results
            html = markdown.markdown(output)
            self.results_text.setHtml(html)
            self.status_bar.showMessage("Analysis complete")

        except Exception as e:
            self.status_bar.showMessage(f"Error: {str(e)}")
            logger.error(f"Error analyzing presentation: {e}")

def main():
    app = QApplication(sys.argv)
    window = PowerPointAnalyzerGUI()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
