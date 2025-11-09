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

## Security Considerations

This application is a local utility and does not interact with networks or external services, which means it is not directly susceptible to many common web-based vulnerabilities like those in the OWASP Top 10. However, users should always:

*   **Backup Important Data:** Before converting critical PowerPoint presentations, ensure they are backed up.
*   **Handle Sensitive Information:** Be mindful of the content within the presentations being converted, especially if they contain sensitive or confidential information. Ensure proper handling and storage of the resulting PDF document.
*   **Keep System and Software Updated:** Ensure your operating system, Microsoft PowerPoint, and Python environment are kept up-to-date with the latest security patches. This is particularly important as the application relies on Microsoft PowerPoint's COM automation.

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
    python3 ppt_to_pdf_converter.py
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
    pyinstaller --onefile --windowed ppt_to_pdf_converter.py
    ```

3.  The final `convert_to_pdf.exe` will be located in the `dist` folder.

---

## License

This project is licensed under the [MIT License](./LICENSE.md) - see the LICENSE.md file for details.

---

## Contributing

Contributions are welcome! If you have suggestions for improvements, bug reports, or want to add new features, please feel free to open an issue or submit a pull request.