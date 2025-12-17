import os
import sys
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox

def main():
    # 1. Hide the main window
    root = tk.Tk()
    root.withdraw()

    # 2. Ask user for the folder
    folder_path = filedialog.askdirectory(title="Select Folder containing PDFs")
    if not folder_path:
        return

    # 3. Create output folder
    output_folder = os.path.join(folder_path, "JPG_Export")
    os.makedirs(output_folder, exist_ok=True)

    # 4. Find PDF files
    files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    if not files:
        messagebox.showerror("Error", "No PDF files found in that folder.")
        return

    # 5. Conversion Settings
    # Zoom 2.0 = ~150 DPI (Medium-High Quality)
    mat = fitz.Matrix(2.0, 2.0) 
    count = 0

    try:
        for filename in files:
            pdf_path = os.path.join(folder_path, filename)
            try:
                doc = fitz.open(pdf_path)
                name_no_ext = os.path.splitext(filename)[0]

                for page_index, page in enumerate(doc):
                    pix = page.get_pixmap(matrix=mat)
                    
                    # Naming: OriginalName_Page01.jpg
                    out_name = f"{name_no_ext}_Page{page_index+1:02d}.jpg"
                    out_path = os.path.join(output_folder, out_name)
                    pix.save(out_path)
                
                count += 1
                doc.close()
            except Exception as e:
                print(f"Skipping bad file: {filename}")

        messagebox.showinfo("Done", f"Converted {count} files.\nSaved in: {output_folder}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

if __name__ == "__main__":
    main()
