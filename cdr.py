import os
import subprocess
import logging
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from reportlab.graphics import renderPM
import tempfile
import shutil

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def convert_cdr_to_ai_free(input_file, output_file):
    """
    Converts a CorelDRAW (.cdr) file to an Adobe Illustrator (.ai) file using free libraries.
    """
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file '{input_file}' does not exist.")
    
    try:
        # Convert CDR to SVG using inkscape (free tool)
        temp_svg = tempfile.NamedTemporaryFile(delete=False, suffix=".svg")
        subprocess.run(["inkscape", input_file, "--export-type=svg", "--export-filename", temp_svg.name], check=True)
        
        # Convert SVG to AI using svglib and reportlab
        drawing = svg2rlg(temp_svg.name)
        with open(output_file, "wb") as ai_file:
            renderPDF.drawToFile(drawing, ai_file)
        
        logging.info(f"Conversion successful: '{output_file}' created.")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error during conversion: {e}")
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
    finally:
        # Clean up temporary SVG file
        if os.path.exists(temp_svg.name):
            os.unlink(temp_svg.name)

# Example usage
if __name__ == "__main__":
    input_cdr = r"C:\Projects\Kal\TEST\Cape Farewell - BM24\Objective - cdr\BM24pt#0.cdr"  # Replace with your .cdr file path
    output_ai = r"C:\Projects\Kal\TEST\Cape Farewell - BM24\Objective - cdr\BM24pt#0.ai"  # Replace with your desired .ai file path
    convert_cdr_to_ai_free(input_cdr, output_ai)
