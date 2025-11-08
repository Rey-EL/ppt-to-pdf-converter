# ppt-to-pdf-converter

A Windows GUI application that automates converting multiple PowerPoint files (`.ppt`, `.pptx`) into PDFs and then merges them into a single document.

---

## **CRITICAL: System Requirements**

This is a **Windows-only** application and requires **Microsoft PowerPoint** to be installed on the system to function. It uses Windows COM automation to programmatically control PowerPoint for high-fidelity conversions.

---

## Key Features

*   **Full GUI Application:** Built with `tkinter` for a user-friendly interface, including a progress bar and live status updates.
*   **PowerPoint Automation:** Uses the `win32com` library to leverage your existing PowerPoint installation for perfect conversions.
*   **Recursive Search:** Scans a selected folder and all its subfolders to find every PowerPoint presentation.
*   **PDF Merging:** After conversion, it uses the `pypdf` library to merge all the newly created PDFs into a single file.
*   **Clean & Safe:** Uses a temporary directory to store intermediate PDFs, which is automatically deleted upon completion, leaving no trace.
*   **Robust Error Handling:** Provides detailed error messages if a file fails to convert or if another issue occurs.

---

## Installation & Setup

1.  **Navigate to the project directory:**
    ```bash
    cd ppt-to-pdf-converter
    ```

2.  **Install dependencies:**
    It's highly recommended to use a virtual environment.
    ```bash
    # Create and activate a virtual environment
    python3 -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`

    # Install the required packages
    pip install -r requirements.txt
    ```

---

## Usage

1.  **Run the script from your terminal:**
    ```bash
    python3 PptToPdfConverter.py
    ```

2.  **Select a Folder:** Click the **"1. Select Folder with Presentations"** button and choose the folder containing your `.ppt` or `.pptx` files.

3.  **Convert and Save:** Click the **"2. Convert and Save As PDF..."** button. A "Save As" dialog will appear for you to name your final merged PDF.

4.  **Monitor Progress:** The application will convert each presentation one by one. The progress bar and status label will update in real-time.

5.  **Done:** A success message will appear when the merge is complete.

---

## Creating a Standalone Executable (Optional)

You can bundle this application into a single `.exe` file using **PyInstaller**. This allows it to be run on other Windows machines (provided they also have PowerPoint installed).

To build the executable:

1.  **Install PyInstaller:**
    ```bash
    pip install pyinstaller
    ```

2.  **Run the build command from the project directory:**
    This command will create a `.spec` file and a `dist` folder containing the final `convert_to_pdf.exe`.
    ```bash
    pyinstaller --onefile --windowed PptToPdfConverter.py
    ```

3.  The final `convert_to_pdf.exe` will be located in the `dist` folder.