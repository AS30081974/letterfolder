import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import PyPDF2
from openpyxl import Workbook, load_workbook
import glob
import os
import threading
from pathlib import Path


class PDFAddressExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Address Extractor")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.pdf_folder_path = tk.StringVar()
        self.excel_file_path = tk.StringVar(value="addresses.xlsx")
        self.extraction_running = False
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF Address Extractor", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # PDF Folder Selection
        ttk.Label(main_frame, text="PDF Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.pdf_folder_path, width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 5))
        ttk.Button(main_frame, text="Browse", 
                  command=self.browse_pdf_folder).grid(row=1, column=2, pady=5)
        
        # Excel File Selection
        ttk.Label(main_frame, text="Excel File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_file_path, width=50).grid(
            row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(10, 5))
        ttk.Button(main_frame, text="Browse", 
                  command=self.browse_excel_file).grid(row=2, column=2, pady=5)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        self.extract_button = ttk.Button(buttons_frame, text="Extract Addresses", 
                                        command=self.extract_addresses_threaded,
                                        style='Accent.TButton')
        self.extract_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_button = ttk.Button(buttons_frame, text="Clear Spreadsheet", 
                                      command=self.clear_spreadsheet)
        self.clear_button.pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Status frame
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
        
        # Status text area
        self.status_text = scrolledtext.ScrolledText(status_frame, height=15, width=70)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure main frame row weights
        main_frame.rowconfigure(5, weight=1)
        
        # Set default folder to current directory
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.pdf_folder_path.set(current_dir)
        
        self.log_message("Application started. Ready to extract addresses from PDFs.")
        
    def browse_pdf_folder(self):
        folder = filedialog.askdirectory(title="Select PDF Folder")
        if folder:
            self.pdf_folder_path.set(folder)
            self.log_message(f"PDF folder selected: {folder}")
    
    def browse_excel_file(self):
        file = filedialog.asksaveasfilename(
            title="Select Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.excel_file_path.set(file)
            self.log_message(f"Excel file selected: {file}")
    
    def log_message(self, message):
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
    
    def extract_text_from_pdfs(self, folder_path):
        """Extract text from PDF files in the specified folder"""
        pdf_texts = {}
        
        # Find all PDF files in the folder
        pdf_pattern = os.path.join(folder_path, "*.pdf")
        pdf_files = glob.glob(pdf_pattern)
        
        if not pdf_files:
            self.log_message("No PDF files found in the selected folder.")
            return pdf_texts
        
        self.log_message(f"Found {len(pdf_files)} PDF files to process.")
        
        for file_path in pdf_files:
            try:
                self.log_message(f"Processing: {os.path.basename(file_path)}")
                
                with open(file_path, 'rb') as pdf_file:
                    reader = PyPDF2.PdfReader(pdf_file)
                    text = ''
                    
                    for page_num in range(len(reader.pages)):
                        page = reader.pages[page_num]
                        text += page.extract_text()
                    
                    # Extract text between ".com" and "Dear"
                    start = text.find("uk_team_gbmailgps@lilly.com")
                    end = text.find("Dear")
                    
                    if start != -1 and end != -1:
                        lines = text[start + 5:end].split('\n')
                        lines = [line.strip() for line in lines if line.strip() != ""]
                        extracted_text = '\n'.join(lines[2:]) if len(lines) > 2 else '\n'.join(lines)
                        pdf_texts[os.path.basename(file_path)] = extracted_text
                        self.log_message(f"Successfully extracted address from {os.path.basename(file_path)}")
                    else:
                        self.log_message(f"Could not find address markers in {os.path.basename(file_path)}")
                        
            except Exception as e:
                self.log_message(f"Error processing {os.path.basename(file_path)}: {str(e)}")
        
        return pdf_texts
    
    def extract_addresses_threaded(self):
        """Run extraction in a separate thread to prevent UI freezing"""
        if self.extraction_running:
            self.log_message("Extraction is already running. Please wait...")
            return
            
        thread = threading.Thread(target=self.extract_addresses)
        thread.daemon = True
        thread.start()
    
    def extract_addresses(self):
        """Main extraction function"""
        try:
            self.extraction_running = True
            self.extract_button.config(state='disabled')
            self.progress.start()
            
            # Validate inputs
            if not self.pdf_folder_path.get():
                messagebox.showerror("Error", "Please select a PDF folder.")
                return
            
            if not self.excel_file_path.get():
                messagebox.showerror("Error", "Please specify an Excel file.")
                return
            
            if not os.path.exists(self.pdf_folder_path.get()):
                messagebox.showerror("Error", "PDF folder does not exist.")
                return
            
            self.log_message("Starting address extraction...")
            
            # Extract text from PDFs
            pdf_texts = self.extract_text_from_pdfs(self.pdf_folder_path.get())
            
            if not pdf_texts:
                self.log_message("No addresses were extracted.")
                return
            
            # Handle Excel file
            excel_path = self.excel_file_path.get()
            
            # Create or load workbook
            if os.path.exists(excel_path):
                try:
                    wb = load_workbook(excel_path)
                    self.log_message(f"Loaded existing Excel file: {excel_path}")
                except Exception as e:
                    self.log_message(f"Error loading Excel file, creating new one: {str(e)}")
                    wb = Workbook()
            else:
                wb = Workbook()
                self.log_message(f"Created new Excel file: {excel_path}")
            
            sheet = wb.active
            if sheet.title == "Sheet":
                sheet.title = "Addresses"
            
            # Add headers if the sheet is empty
            if sheet.max_row == 1 and sheet['A1'].value is None:
                sheet['A1'] = "PDF File"
                sheet['B1'] = "Extracted Address"
                current_row = 2
            else:
                current_row = sheet.max_row + 1
            
            # Write extracted addresses to Excel
            self.log_message("Writing addresses to Excel...")
            
            for pdf_file, address in pdf_texts.items():
                sheet.cell(row=current_row, column=1).value = pdf_file
                sheet.cell(row=current_row, column=2).value = address
                current_row += 1
            
            # Auto-adjust column widths
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                sheet.column_dimensions[column_letter].width = adjusted_width
            
            # Save the workbook
            wb.save(excel_path)
            wb.close()
            
            self.log_message(f"Successfully extracted {len(pdf_texts)} addresses and saved to {excel_path}")
            messagebox.showinfo("Success", f"Extracted {len(pdf_texts)} addresses successfully!")
            
        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Error", error_msg)
        
        finally:
            self.extraction_running = False
            self.extract_button.config(state='normal')
            self.progress.stop()
    
    def clear_spreadsheet(self):
        """Clear the contents of the Excel spreadsheet"""
        try:
            excel_path = self.excel_file_path.get()
            
            if not excel_path:
                messagebox.showerror("Error", "Please specify an Excel file.")
                return
            
            # Confirm action
            result = messagebox.askyesno(
                "Confirm Clear", 
                f"Are you sure you want to clear all data from {os.path.basename(excel_path)}?\n\nThis action cannot be undone."
            )
            
            if not result:
                return
            
            # Create a new workbook or clear existing one
            wb = Workbook()
            sheet = wb.active
            sheet.title = "Addresses"
            
            # Add headers
            sheet['A1'] = "PDF File"
            sheet['B1'] = "Extracted Address"
            
            # Save the workbook
            wb.save(excel_path)
            wb.close()
            
            self.log_message(f"Cleared spreadsheet: {excel_path}")
            messagebox.showinfo("Success", "Spreadsheet cleared successfully!")
            
        except Exception as e:
            error_msg = f"Error clearing spreadsheet: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Error", error_msg)


def main():
    root = tk.Tk()
    
    # Configure ttk styles for a modern look
    style = ttk.Style()
    style.theme_use('clam')  # Use a modern theme
    
    # Create custom style for accent button
    style.configure('Accent.TButton', foreground='white', background='#0078d4')
    style.map('Accent.TButton', 
             background=[('active', '#106ebe'), ('pressed', '#005a9e')])
    
    app = PDFAddressExtractorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
