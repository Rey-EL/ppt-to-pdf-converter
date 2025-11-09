# Copyright (c) 2025 Elvis Reyes
#
# This software is licensed under the MIT License.
# See the LICENSE.md file in the project root for full license terms.

import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import pythoncom
from pypdf import PdfWriter
import traceback
import tempfile

def convert_ppt_to_pdf(source_ppt_path, dest_pdf_path):
    """Uses PowerPoint COM automation to convert a .ppt(x) to a .pdf."""
    pythoncom.CoInitialize()
    powerpoint = None
    try:
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        deck = powerpoint.Presentations.Open(source_ppt_path, WithWindow=False)
        deck.SaveAs(dest_pdf_path, 32)
        deck.Close()
        return True
    except Exception:
        raise
    finally:
        if powerpoint:
            powerpoint.Quit()
        pythoncom.CoUninitialize()

def main_process(source_folder, final_pdf_path, app_instance):
    """Main logic that finds, converts, and merges files, updating the GUI."""
    ppt_files = []
    valid_extensions = ('.pptx', '.ppt')
    for root, _, files in os.walk(source_folder):
        for file in files:
            if file.lower().endswith(valid_extensions):
                ppt_files.append(os.path.join(root, file))

    if not ppt_files:
        return "No PowerPoint files found in the selected folder."

    ppt_files.sort()
    
    # --- THIS IS THE UPDATED PART ---
    # Create a secure, temporary directory that cleans up automatically
    with tempfile.TemporaryDirectory() as temp_pdf_folder:
        app_instance.progress_bar['maximum'] = len(ppt_files)
        converted_pdfs = []

        # Step 1: Convert all PowerPoints to PDF inside the temp folder
        for i, ppt_path in enumerate(ppt_files):
            app_instance.update_status(f"Converting: {os.path.basename(ppt_path)}")
            pdf_name = os.path.splitext(os.path.basename(ppt_path))[0] + ".pdf.pdf"
            dest_pdf = os.path.join(temp_pdf_folder, pdf_name)
            if convert_ppt_to_pdf(os.path.abspath(ppt_path), os.path.abspath(dest_pdf)):
                converted_pdfs.append(dest_pdf)
            app_instance.progress_bar['value'] = i + 1
            app_instance.root.update_idletasks()

        if not converted_pdfs:
            return "Conversion failed. No PDFs were created."

        # Step 2: Merge the newly created PDFs
        app_instance.update_status("Merging all PDFs...")
        pdf_writer = PdfWriter()
        for pdf_path in converted_pdfs:
            pdf_writer.append(pdf_path)

        # Step 3: Save the final PDF
        with open(final_pdf_path, "wb") as out_file:
            pdf_writer.write(out_file)
        pdf_writer.close()
    # --- END OF UPDATED PART (the temp folder is auto-deleted here) ---
    
    return f"Success! {len(converted_pdfs)} presentations converted and merged."

# --- GUI Application Code (No changes needed here) ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT to PDF Converter")
        self.root.geometry("500x200")

        self.folder_label = tk.Label(root, text="No folder selected.", wraplength=480)
        self.folder_label.pack(pady=10)
        
        self.select_button = tk.Button(root, text="1. Select Folder with Presentations", command=self.select_folder)
        self.select_button.pack(pady=5)

        self.merge_button = tk.Button(root, text="2. Convert and Save As PDF...", command=self.run_process)
        self.merge_button.pack(pady=10)
        
        self.status_label = tk.Label(root, text="Waiting to start...")
        self.status_label.pack(pady=5)
        
        self.progress_bar = ttk.Progressbar(root, orient='horizontal', length=400, mode='determinate')
        self.progress_bar.pack(pady=5)
        
        self.selected_folder = ""

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_folder = folder
            self.folder_label.config(text=f"Selected: {folder}")

    def run_process(self):
        if not self.selected_folder:
            messagebox.showerror("Error", "Please select a folder first.")
            return
        
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save Final PDF As...")
        if not output_path: return

        try:
            self.progress_bar['value'] = 0
            self.update_status("Starting...")
            self.merge_button.config(state="disabled")
            
            status = main_process(self.selected_folder, output_path, self)
            messagebox.showinfo("Result", status)

        except Exception as e:
            error_details = traceback.format_exc()
            messagebox.showerror("A Critical Error Occurred", f"The application had to stop.\n\nERROR: {str(e)}\n\nDETAILS:\n{error_details}")
        finally:
            self.merge_button.config(state="normal")
            self.update_status("Done. Ready for next task.")
            self.progress_bar['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
