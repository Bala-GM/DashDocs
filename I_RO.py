'''import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pdfplumber

class PDFtoExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        
        self.upload_button = tk.Button(root, text="Upload PDF File", command=self.upload_file)
        self.upload_button.pack(pady=10)
        
        self.process_button = tk.Button(root, text="Convert to Excel", command=self.process_file)
        self.process_button.pack(pady=10)
        
        self.file_path = None
    
    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.file_path = file_path
            messagebox.showinfo("File Upload", f"Uploaded file: {self.file_path}")
    
    def process_file(self):
        if not self.file_path:
            messagebox.showwarning("No File", "No file has been uploaded.")
            return
        
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            self.convert_pdf_to_excel(self.file_path, output_path)
            messagebox.showinfo("File Saved", f"Excel file saved to: {output_path}")

    def convert_pdf_to_excel(self, pdf_path, excel_path):
        with pdfplumber.open(pdf_path) as pdf:
            with pd.ExcelWriter(excel_path) as writer:
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for j, table in enumerate(tables):
                        df = pd.DataFrame(table[1:], columns=table[0])
                        df.to_excel(writer, index=False, sheet_name=f"Page_{i+1}_Table_{j+1}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoExcelApp(root)
    root.mainloop()'''


'''Certainly! There are several libraries and frameworks available for working with PDFs in various programming languages. Here are some popular ones:

### Python

1. **pdfplumber**: A Python library for extracting information from PDF documents. It focuses on text extraction and table parsing from PDFs.

2. **PyPDF2**: A pure-Python library built as a PDF toolkit. It can extract document information, split and merge PDFs, encrypt and decrypt PDF files, and more.

3. **ReportLab**: A Python library for creating PDF documents programmatically. It allows for the creation of complex documents with charts, tables, and vector graphics.

4. **PyMuPDF (fitz)**: A Python binding for MuPDF, a lightweight PDF, XPS, and E-book viewer. It provides both text extraction and high-level object access to PDF files.

5. **Camelot**: A Python library that allows you to extract tables from PDFs into pandas DataFrame format. It's built on top of pdfminer.six and uses a combination of image processing techniques and OCR to detect and extract tables.

### Java

1. **iText**: A Java library that allows developers to create, manipulate, and extract data from PDF files. It supports PDF generation, text extraction, and manipulation of PDF documents.

2. **Apache PDFBox**: An open-source Java library for working with PDF documents. It provides functionalities for creation, manipulation (merging, splitting), text extraction, and more.

3. **PDF.js**: Although primarily a JavaScript library for rendering PDFs in the browser, it also has Java bindings (PDF.js Express) that can be used for server-side PDF manipulation.

### JavaScript/Node.js

1. **pdf-lib**: A JavaScript library for creating and manipulating PDF files. It supports PDF generation, merging, splitting, encryption, and text extraction.

2. **pdfkit**: A Node.js library that generates PDFs using HTML and CSS as input. It allows for the creation of complex PDF documents with embedded fonts, images, and vector graphics.

3. **pdf-parse**: A Node.js library that extracts text and metadata from PDF files using a straightforward API.

### C#

1. **iTextSharp**: A C# port of the iText library, providing similar functionalities for PDF generation, manipulation, and extraction.

2. **PDFsharp**: A .NET library for creating and processing PDF documents. It allows for drawing graphics, text, and images in PDF files.

3. **Docotic.Pdf**: A .NET library that supports PDF manipulation, text extraction, digital signatures, and more.

### Ruby

1. **Prawn**: A pure Ruby library for generating PDF documents programmatically. It allows for embedding text, graphics, and images in PDFs.

2. **PDF::Reader**: A Ruby library for reading content from PDF files. It provides APIs to extract text and metadata from PDF documents.

These libraries vary in terms of features, ease of use, and language support. Depending on your specific requirements (such as text extraction, document generation, or advanced manipulation), you can choose the appropriate library that best fits your project needs.'''


import os
print(os.environ.get('JAVA_HOME'))

import subprocess

try:
    result = subprocess.run(['java', '-version'], capture_output=True, text=True)
    print(result.stdout)
except FileNotFoundError:
    print("Java command not found. Ensure JAVA_HOME is set correctly.")

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import tabula

class PDFtoExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        
        self.upload_button = tk.Button(root, text="Upload PDF File", command=self.upload_file)
        self.upload_button.pack(pady=10)
        
        self.process_button = tk.Button(root, text="Convert to Excel", command=self.process_file)
        self.process_button.pack(pady=10)
        
        self.file_path = None
    
    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.file_path = file_path
            messagebox.showinfo("File Upload", f"Uploaded file: {self.file_path}")
    
    def process_file(self):
        if not self.file_path:
            messagebox.showwarning("No File", "No file has been uploaded.")
            return
        
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            self.convert_pdf_to_excel(self.file_path, output_path)
            messagebox.showinfo("File Saved", f"Excel file saved to: {output_path}")

    def convert_pdf_to_excel(self, pdf_path, excel_path):
        # Extract tables using Tabula-py
        tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
        
        # Write each table to a separate sheet in the Excel file
        with pd.ExcelWriter(excel_path) as writer:
            for i, table in enumerate(tables):
                df = pd.DataFrame(table)
                df.to_excel(writer, index=False, sheet_name=f"Page_{i+1}_Table_{i+1}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoExcelApp(root)
    root.mainloop()
