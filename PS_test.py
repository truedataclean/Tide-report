import os
import time
import subprocess

def get_file_info(file_path):
    try:
        # Check if the file exists
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return

        # Get file size
        file_size = os.path.getsize(file_path)

        # Get file creation and modification times
        creation_time = os.path.getctime(file_path)
        modification_time = os.path.getmtime(file_path)

        print(f"File: {file_path}")
        print(f"Size: {file_size} bytes")
        print(f"Created: {time.ctime(creation_time)}")
        print(f"Last Modified: {time.ctime(modification_time)}")

    except Exception as e:
        print(f"An error occurred: {e}")

def convert_ps_to_pdf(ps_file_path, pdf_file_path):
    try:
        # Use Ghostscript to convert PS to PDF
        command = [
            "gswin64c.exe",  # Path to Ghostscript executable, adjust if necessary
            "-dBATCH",
            "-dNOPAUSE",
            "-sDEVICE=pdfwrite",
            f"-sOutputFile={pdf_file_path}",
            ps_file_path
        ]
        subprocess.run(command, check=True, shell=True)
        print(f"Converted {ps_file_path} to {pdf_file_path}")
    except Exception as e:
        print(f"An error occurred during conversion: {e}")

if __name__ == "__main__":
    # Replace with the path to your PS file
    ps_file_path = r"C:\Projects\Kal\TEST\Cape Farewell - BM24\WIN11 - ps\BM24pt#0.ps"
    pdf_file_path = r"C:\Projects\Kal\TEST\Cape Farewell - BM24\WIN11 - ps\BM24pt#0.pdf"

    # Get file info
    get_file_info(ps_file_path)

    # Convert PS to PDF
    convert_ps_to_pdf(ps_file_path, pdf_file_path)