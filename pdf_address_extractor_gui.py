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
        self.root.title("� Address Extractor Pro")
        self.root.geometry("900x700")
        self.root.configure(bg='#f8f9fa')
        self.root.resizable(True, True)
        
        # Set window icon and styling
        try:
            # Try to set a custom icon if available
            icon_path = os.path.join(os.path.dirname(__file__), "app_icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                # Create a simple text-based icon using tkinter's built-in method
                self.root.iconname("AddressExtractor")
        except:
            pass
        
        # Variables
        self.pdf_folder_path = tk.StringVar()
        self.excel_file_path = tk.StringVar(value="addresses.xlsx")
        self.extraction_running = False
        
        self.setup_modern_styles()
        self.setup_ui()
        
    def setup_modern_styles(self):
        """Configure modern ttk styles"""
        self.style = ttk.Style()
        
        # Use a modern theme as base
        self.style.theme_use('clam')
        
        # Configure modern colors
        colors = {
            'primary': '#2563eb',      # Blue
            'primary_dark': '#1d4ed8',
            'secondary': '#64748b',    # Slate
            'success': '#059669',      # Green
            'danger': '#dc2626',       # Red
            'warning': '#d97706',      # Orange
            'light': '#f8fafc',        # Light gray
            'dark': '#1e293b',         # Dark gray
            'white': '#ffffff',
            'border': '#e2e8f0'
        }
        
        # Configure button styles
        self.style.configure('Primary.TButton',
                           background=colors['primary'],
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none',
                           padding=(20, 12))
        
        self.style.map('Primary.TButton',
                      background=[('active', colors['primary_dark']),
                                ('pressed', colors['primary_dark'])])
        
        self.style.configure('Secondary.TButton',
                           background=colors['secondary'],
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none',
                           padding=(16, 10))
        
        self.style.map('Secondary.TButton',
                      background=[('active', colors['dark']),
                                ('pressed', colors['dark'])])
        
        self.style.configure('Success.TButton',
                           background=colors['success'],
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none',
                           padding=(16, 10))
        
        self.style.map('Success.TButton',
                      background=[('active', '#047857'),
                                ('pressed', '#047857')])
        
        # Configure frame styles
        self.style.configure('Card.TFrame',
                           background='white',
                           relief='flat',
                           borderwidth=1)
        
        self.style.configure('Header.TFrame',
                           background=colors['primary'],
                           relief='flat')
        
        # Configure label styles
        self.style.configure('Title.TLabel',
                           background='white',
                           foreground=colors['dark'],
                           font=('Segoe UI', 24, 'bold'))
        
        self.style.configure('Subtitle.TLabel',
                           background='white',
                           foreground=colors['secondary'],
                           font=('Segoe UI', 11))
        
        self.style.configure('CardTitle.TLabel',
                           background='white',
                           foreground=colors['dark'],
                           font=('Segoe UI', 12, 'bold'))
        
        self.style.configure('FieldLabel.TLabel',
                           background='white',
                           foreground=colors['dark'],
                           font=('Segoe UI', 10))
        
        # Configure entry styles
        self.style.configure('Modern.TEntry',
                           fieldbackground='white',
                           borderwidth=2,
                           relief='solid',
                           insertcolor=colors['primary'],
                           padding=(12, 8))
        
        # Configure progressbar styles
        self.style.configure('Modern.Horizontal.TProgressbar',
                           background=colors['primary'],
                           troughcolor='#e5e7eb',
                           borderwidth=0,
                           lightcolor=colors['primary'],
                           darkcolor=colors['primary'])
        
        # Configure labelframe styles
        self.style.configure('Modern.TLabelframe',
                           background='white',
                           relief='flat',
                           borderwidth=0)
        
        self.style.configure('Modern.TLabelframe.Label',
                           background='white',
                           foreground=colors['dark'],
                           font=('Segoe UI', 12, 'bold'))
        
    def setup_ui(self):
        # Create main container with padding
        main_container = tk.Frame(self.root, bg='#f8f9fa')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Configure grid weights
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=1)
        
        # Header section
        self.create_header(main_container)
        
        # Main content area
        content_frame = tk.Frame(main_container, bg='#f8f9fa')
        content_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        content_frame.grid_rowconfigure(2, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)
        
        # File selection card
        self.create_file_selection_card(content_frame)
        
        # Action buttons card
        self.create_action_buttons_card(content_frame)
        
        # Status and progress card
        self.create_status_card(content_frame)
        
        # Set default folder
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.pdf_folder_path.set(current_dir)
        
        self.log_message("🚀 Address Extractor Pro started. Ready to extract addresses from PDFs.")
        
    def create_header(self, parent):
        """Create modern header section"""
        header_frame = ttk.Frame(parent, style='Card.TFrame')
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        header_frame.grid_columnconfigure(0, weight=1)
        
        # Add padding inside the card
        header_content = tk.Frame(header_frame, bg='white')
        header_content.pack(fill=tk.BOTH, expand=True, padx=30, pady=25)
        
        # Title with icon
        title_frame = tk.Frame(header_content, bg='white')
        title_frame.pack(anchor=tk.W)
        
        title_label = ttk.Label(title_frame, text="🏢📋 Address Extractor Pro", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        # Subtitle
        subtitle_label = ttk.Label(header_content, 
                                 text="Extract addresses from PDF files and save to Excel spreadsheets", 
                                 style='Subtitle.TLabel')
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))
        
    def create_file_selection_card(self, parent):
        """Create file selection card"""
        card_frame = ttk.Frame(parent, style='Card.TFrame')
        card_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        card_frame.grid_columnconfigure(1, weight=1)
        
        # Card content with padding
        content = tk.Frame(card_frame, bg='white')
        content.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        content.grid_columnconfigure(1, weight=1)
        
        # Card title
        ttk.Label(content, text="📁 File Selection", style='CardTitle.TLabel').grid(
            row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 20))
        
        # PDF Folder Selection
        ttk.Label(content, text="PDF Folder:", style='FieldLabel.TLabel').grid(
            row=1, column=0, sticky=tk.W, pady=(0, 15), padx=(0, 15))
        
        pdf_entry = ttk.Entry(content, textvariable=self.pdf_folder_path, 
                             style='Modern.TEntry', font=('Segoe UI', 10))
        pdf_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(0, 15), padx=(0, 10))
        
        pdf_browse_btn = ttk.Button(content, text="📂 Browse", 
                                   command=self.browse_pdf_folder, style='Secondary.TButton')
        pdf_browse_btn.grid(row=1, column=2, pady=(0, 15))
        
        # Excel File Selection
        ttk.Label(content, text="Excel File:", style='FieldLabel.TLabel').grid(
            row=2, column=0, sticky=tk.W, padx=(0, 15))
        
        excel_entry = ttk.Entry(content, textvariable=self.excel_file_path, 
                               style='Modern.TEntry', font=('Segoe UI', 10))
        excel_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        excel_browse_btn = ttk.Button(content, text="📊 Browse", 
                                     command=self.browse_excel_file, style='Secondary.TButton')
        excel_browse_btn.grid(row=2, column=2)
        
    def create_action_buttons_card(self, parent):
        """Create action buttons card"""
        card_frame = ttk.Frame(parent, style='Card.TFrame')
        card_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Card content with padding
        content = tk.Frame(card_frame, bg='white')
        content.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        # Card title
        ttk.Label(content, text="⚡ Actions", style='CardTitle.TLabel').pack(
            anchor=tk.W, pady=(0, 20))
        
        # Buttons container
        buttons_frame = tk.Frame(content, bg='white')
        buttons_frame.pack(anchor=tk.W)
        
        # Extract button with icon
        self.extract_button = ttk.Button(buttons_frame, text="🔍 Extract Addresses", 
                                        command=self.extract_addresses_threaded,
                                        style='Primary.TButton')
        self.extract_button.pack(side=tk.LEFT, padx=(0, 15))
        
        # Clear button
        self.clear_button = ttk.Button(buttons_frame, text="🧹 Clear Spreadsheet", 
                                      command=self.clear_spreadsheet,
                                      style='Secondary.TButton')
        self.clear_button.pack(side=tk.LEFT)
        
        # Progress bar (initially hidden)
        self.progress_frame = tk.Frame(content, bg='white')
        self.progress_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Label(self.progress_frame, text="Processing...", 
                 style='FieldLabel.TLabel').pack(anchor=tk.W, pady=(0, 5))
        
        self.progress = ttk.Progressbar(self.progress_frame, mode='indeterminate', 
                                       style='Modern.Horizontal.TProgressbar')
        self.progress.pack(fill=tk.X)
        
        # Hide progress initially
        self.progress_frame.pack_forget()
        
    def create_status_card(self, parent):
        """Create status and log card"""
        card_frame = ttk.Frame(parent, style='Card.TFrame')
        card_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        card_frame.grid_rowconfigure(1, weight=1)
        card_frame.grid_columnconfigure(0, weight=1)
        
        # Card content with padding
        content = tk.Frame(card_frame, bg='white')
        content.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        content.grid_rowconfigure(1, weight=1)
        content.grid_columnconfigure(0, weight=1)
        
        # Card title
        ttk.Label(content, text="📋 Status Log", style='CardTitle.TLabel').grid(
            row=0, column=0, sticky=tk.W, pady=(0, 15))
        
        # Status text area with modern styling
        text_frame = tk.Frame(content, bg='#f8fafc', relief='solid', bd=1)
        text_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)
        
        self.status_text = scrolledtext.ScrolledText(
            text_frame, 
            height=12, 
            width=70,
            bg='#f8fafc',
            fg='#334155',
            font=('Consolas', 10),
            relief='flat',
            borderwidth=0,
            insertbackground='#2563eb',
            selectbackground='#dbeafe',
            wrap=tk.WORD
        )
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=15, pady=15)
        
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
            self.progress_frame.pack(fill=tk.X, pady=(20, 0))
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
            self.progress_frame.pack_forget()
    
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
