# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "markdown",
#     "matplotlib",
#     "python-pptx",
#     "PyQt6",
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
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QFileDialog, QLineEdit, 
                            QTextEdit, QCheckBox, QGroupBox, QStatusBar)
from typing import Dict, List, Set, Tuple


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

def format_font_report(font_usage: Dict[int, Dict[str, Set[str]]], 
                     all_fonts: Set[str],
                     system_fonts: Set[str]) -> str:
    """Create a formatted report showing which slides use each font."""
    result = ""

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

    result += "## Font Usage\n"
    
    if font_to_slides:
        for font in sorted(font_to_slides.keys()):
            if font:  # Skip None values
                status = "" if font.lower() in system_fonts else " **(missing)**"
                # Convert slide numbers to a readable string
                slides = sorted(font_to_slides[font])
                slides_str = ", ".join(str(slide) for slide in slides)
        
                result += f"### {font}{status}\nSlides: {slides_str}\n"
    else:
        result += "(no fonts used in presentation)\n"
    
    # Print summary statistics
    total_fonts = len(font_to_slides)
    missing_fonts = sum(1 for font in font_to_slides if font.lower() not in system_fonts)
    
    result += "### Font Summary\n"
    result += f"Total unique fonts: {total_fonts}<br />\n"
    result += f"Missing fonts: {missing_fonts}\n"
    result += "***\n"

    return result

def generate_font_report(pptx_path: str) -> str:
    # Get system fonts
    system_fonts = get_system_fonts()
    
    # Analyze presentation
    font_usage, all_fonts = analyze_fonts(pptx_path)
    
    # Print report
    return format_font_report(font_usage, all_fonts, system_fonts)

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
