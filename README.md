# PowerPoint to PDF Converter & Merger

This is a complete GUI application for Windows that automates the process of converting all PowerPoint (.ppt and .pptx) files within a folder into PDFs and then merging them into a single, combined PDF document.

## Features

*   **Full GUI Application:** Built with `tkinter`, providing a user-friendly interface with folder selection, a "Save As" dialog, a live status label, and a progress bar.
*   **PowerPoint Automation:** Uses the `win32com.client` library to programmatically open Microsoft PowerPoint in the background, convert presentations to high-fidelity PDFs, and then close.
*   **PDF Merging:** After conversion, the script uses the `pypdf` library to merge all the newly created PDFs into a single file.
*   **Safe Temp File Handling:** Creates a secure temporary directory (using `tempfile`) to store the intermediate PDFs. This directory is automatically deleted after the final merge is complete, leaving no messy files behind.
*   **Recursive Search:** Scans the selected folder and all its subfolders to find every PowerPoint file, ensuring none are missed.
*   **Error Handling:** Includes `try...except` blocks to provide detailed error messages to the user if a file fails to convert or the application encounters a critical error.

## Deployment / Executable

A `convert_to_pdf.spec` file is included in this repository. This file is a "recipe" for PyInstaller, which compiles this Python script into a standalone `convert_to_pdf.exe` for Windows. The `console=False` flag ensures it runs as a true windowed application without a background terminal.

## How to Use

1.  Run the script or the compiled `.exe` file.
2.  Click the "1. Select Folder with Presentations" button and choose your folder.
3.  Click the "2. Convert and Save As PDF..." button to choose where to save your new file.
4.  The progress bar and status label will update as the script works.
5.  A success message will appear when the merge is complete.
