"""
GUI Components for Document Processor Pro
Modern interface components using tkinter and ttk.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import glob
from typing import Callable, List, Any


class ModernStyles:
    """Modern styling configuration for the GUI with contemporary aesthetics."""
    
    @staticmethod
    def configure_styles():
        """Configure contemporary ttk styles with modern design principles."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Modern color scheme
        colors = {
            'primary': '#3b82f6',          # Brighter blue
            'primary_dark': '#2563eb',
            'primary_light': '#93c5fd',
            'secondary': "#434953",
            'secondary_light': '#cbd5e1',
            'success': '#10b981',          # Vibrant green
            'warning': '#f59e0b',
            'danger': '#ef4444',            # Deeper red
            'light': "#f8f8fc",
            'dark': '#0f172a',             # Deeper dark
            'background': '#ffffff',
            'card_bg': '#f3f4f6',  # Light gray for cards
            'border': "#adadad"
        }
        
        # Configure modern button styles with rounded corners
        for btn_style in ['Primary', 'Success', 'Secondary']:
            base_color = colors[btn_style.lower()]
            hover_color = colors[f'{btn_style.lower()}_dark'] if f'{btn_style.lower()}_dark' in colors else '#334155'
            
            style.configure(f'{btn_style}.TButton',
                           background=base_color,
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none',
                           padding=(16, 12),
                           relief='flat')
            
            style.map(f'{btn_style}.TButton',
                     background=[('active', hover_color),
                               ('pressed', hover_color)],
                     relief=[('active', 'flat'),
                            ('pressed', 'flat')])
        
        # Modern frame styles with subtle borders
        style.configure('Card.TFrame',
                       relief='solid',
                       borderwidth=5,
                       bordercolor=colors['border'])
        
        style.configure('Header.TFrame',
                       relief='flat')
        
        style.configure('Title.TLabel',        # Add white background
               foreground=colors['dark'],
               font=('Segoe UI', 24, 'bold'))

        style.configure('Subtitle.TLabel',       # Add white background
                    foreground=colors['secondary'],
                    font=('Segoe UI', 12))

        style.configure('CardTitle.TLabel',       # Add white background
                    foreground=colors['dark'],
                    font=('Segoe UI', 15, 'bold'))

        style.configure('Body.TLabel',        # Add white background
                    foreground=colors['secondary'],
                    font=('Segoe UI', 10))
        
        

        
        # Modern progressbar with smoother appearance
        style.configure('TProgressbar',
                       background=colors['primary'],
                       troughcolor=colors['secondary_light'],
                       borderwidth=0,
                       thickness=6,
                       lightcolor=colors['primary'],
                       darkcolor=colors['primary'])
        
        # Configure scrollbar for a cleaner look
        style.configure('TScrollbar',
                       background=colors['light'],
                       borderwidth=0,
                       arrowsize=14,
                       relief='flat',
                       troughcolor=colors['light'])
        
        style.map('TScrollbar',
                 background=[('active', colors['secondary_light']),
                           ('pressed', colors['secondary_light'])])


class FileSelectionCard:
    """File selection card component."""
    
    def __init__(self, parent, title: str, file_types: List[tuple], 
                 on_selection_change: Callable = None):
        self.parent = parent
        self.title = title
        self.file_types = file_types
        self.on_selection_change = on_selection_change
        
        # Create main frame
        self.frame = ttk.Frame(parent, style='Card.TFrame')
        self.selected_files = []
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create the file selection widgets matching original design."""
        # Card header
        header_frame = ttk.Frame(self.frame)
        header_frame.pack(fill='x', padx=20, pady=(20, 10))
        
        title_label = ttk.Label(header_frame, text=self.title, style='CardTitle.TLabel')
        title_label.pack(side='left')
        
        # File info and buttons row
        info_button_frame = ttk.Frame(self.frame)
        info_button_frame.pack(fill='x', padx=20, pady=(0, 10))
        
        self.info_label = ttk.Label(info_button_frame, text="No folder selected", 
                                   style='Body.TLabel')
        self.info_label.pack(side='left')
        
        # Buttons on the right
        button_frame = ttk.Frame(info_button_frame)
        button_frame.pack(side='right')
        
        self.select_btn = ttk.Button(button_frame, text="📁 Select Folder",
                                    command=self._select_folder,
                                    style='Primary.TButton')
        self.select_btn.pack(side='left', padx=(0, 8))
        
        self.clear_btn = ttk.Button(button_frame, text="🗑️ Clear",
                                   command=self._clear_selection,
                                   style='Secondary.TButton')
        self.clear_btn.pack(side='left')
        
        # File list with scrollbar
        list_frame = ttk.Frame(self.frame)
        list_frame.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # Create listbox with original styling
        self.listbox = tk.Listbox(list_frame, 
                                 height=12,
                                 font=('Segoe UI', 9),
                                 selectmode='extended',
                                 bg='white',
                                 fg='#1e293b',
                                 selectbackground='#2563eb',
                                 selectforeground='white',
                                 borderwidth=1,
                                 relief='solid')
        
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical',
                                 command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        
        self.listbox.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
    
    def _select_folder(self):
        """Open folder dialog to select a folder and find PDF files within it."""
        folder = filedialog.askdirectory(
            title="Select Folder Containing PDF Files"
        )
        
        if folder:
            # Find all PDF files in the selected folder and subfolders
            pdf_files = []
            
            # Search for PDF files recursively
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.pdf'):
                        pdf_files.append(os.path.join(root, file))
            
            if pdf_files:
                self.selected_files = sorted(pdf_files)
                self._update_display()
                
                if self.on_selection_change:
                    self.on_selection_change(self.selected_files)
            else:
                messagebox.showinfo(
                    "No PDF Files", 
                    f"No PDF files found in the selected folder:\n{folder}"
                )
    
    def _select_files(self):
        """Open file dialog to select files."""
        files = filedialog.askopenfilenames(
            title=f"Select {self.title}",
            filetypes=self.file_types
        )
        
        if files:
            self.selected_files = list(files)
            self._update_display()
            
            if self.on_selection_change:
                self.on_selection_change(self.selected_files)
    
    def _clear_selection(self):
        """Clear selected files."""
        self.selected_files = []
        self._update_display()
        
        if self.on_selection_change:
            self.on_selection_change(self.selected_files)
    
    def _update_display(self):
        """Update the display with selected files."""
        self.listbox.delete(0, tk.END)
        
        if self.selected_files:
            # Determine if files are from the same folder
            if len(self.selected_files) > 1:
                common_folder = os.path.commonpath(self.selected_files)
                if common_folder and len(common_folder) > 3:  # More than just drive letter
                    folder_name = os.path.basename(common_folder)
                    self.info_label.config(text=f"{len(self.selected_files)} PDF files from '{folder_name}' folder")
                else:
                    self.info_label.config(text=f"{len(self.selected_files)} PDF files selected")
            else:
                self.info_label.config(text=f"{len(self.selected_files)} file selected")
            
            # Show filenames in the listbox
            for file_path in self.selected_files:
                filename = os.path.basename(file_path)
                # If files are from the same folder, show just filename, otherwise show relative path
                if len(self.selected_files) > 1:
                    try:
                        common_folder = os.path.commonpath(self.selected_files)
                        if common_folder:
                            relative_path = os.path.relpath(file_path, common_folder)
                            display_name = relative_path if '\\' in relative_path or '/' in relative_path else filename
                        else:
                            display_name = filename
                    except:
                        display_name = filename
                else:
                    display_name = filename
                
                self.listbox.insert(tk.END, display_name)
        else:
            self.info_label.config(text="No folder selected")
    
    def get_files(self) -> List[str]:
        """Get the selected files."""
        return self.selected_files
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class ActionCard:
    """Action card with buttons."""
    
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent, style='Card.TFrame')
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create action widgets matching original design."""
        # Card header
        header_frame = ttk.Frame(self.frame)
        header_frame.pack(fill='x', padx=20, pady=(20, 10))
        
        title_label = ttk.Label(header_frame, text="🎯 Actions", style='CardTitle.TLabel')
        title_label.pack(side='left')
        
        # Button grid layout
        button_container = ttk.Frame(self.frame)
        button_container.pack(fill='x', padx=20, pady=(0, 20))
        
        # Top row of buttons
        top_row = ttk.Frame(button_container)
        top_row.pack(fill='x', pady=(0, 10))
        
        self.extract_btn = ttk.Button(top_row, 
                                     text="🔍 Extract Addresses",
                                     style='Success.TButton')
        self.extract_btn.pack(side='left', fill='x', expand=True, padx=(0, 8))
        
        self.print_btn = ttk.Button(top_row,
                                   text="🖨️ Print PDFs",
                                   style='Primary.TButton')
        self.print_btn.pack(side='right', fill='x', expand=True, padx=(8, 0))
        
        # Bottom row with stop button
        bottom_row = ttk.Frame(button_container)
        bottom_row.pack(fill='x')

        self.stop_btn = ttk.Button(bottom_row,
                                  text="⏹️ Stop Processing",
                                  style='Secondary.TButton',
                                  state='disabled')
        self.stop_btn.pack(fill='x')
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class LoggingCard:
    """Logging card component."""
    
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent, style='Card.TFrame')
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create logging widgets matching original design."""
        # Card header
        header_frame = ttk.Frame(self.frame)
        header_frame.pack(fill='x', padx=20, pady=(20, 10))
        
        title_label = ttk.Label(header_frame, text="📋 Processing Log", 
                               style='CardTitle.TLabel')
        title_label.pack(side='left', padx=(0, 10))
        
        # Clear button in header
        self.clear_btn = ttk.Button(header_frame, text="🗑️ Clear",
                                   style='Secondary.TButton')
        self.clear_btn.pack(side='right')
        
        # Progress bar with original styling
        progress_frame = ttk.Frame(self.frame)
        progress_frame.pack(fill='x', padx=20, pady=(0, 10))
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress.pack(fill='x')
        
        # Log text area with original styling
        log_frame = ttk.Frame(self.frame)
        log_frame.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        self.log_text = tk.Text(log_frame, 
                               height=15, 
                               wrap=tk.WORD,
                               font=('Consolas', 9),
                               bg='#f8fafc', 
                               fg='#1e293b',
                               selectbackground='#2563eb',
                               selectforeground='white',
                               borderwidth=1,
                               relief='solid',
                               padx=10,
                               pady=8)
        
        log_scrollbar = ttk.Scrollbar(log_frame, orient='vertical',
                                     command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scrollbar.pack(side='right', fill='y')
    
    def log(self, message: str):
        """Add a message to the log."""
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.parent.update_idletasks()
    
    def clear_log(self):
        """Clear the log."""
        self.log_text.delete(1.0, tk.END)
    
    def show_progress(self):
        """Show progress bar."""
        self.progress.start(10)
    
    def hide_progress(self):
        """Hide progress bar."""
        self.progress.stop()
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class ExcelSelectionCard:
    """Excel file selection card component."""
    
    def __init__(self, parent, on_selection_change: Callable = None):
        self.parent = parent
        self.on_selection_change = on_selection_change
        
        # Create main frame
        self.frame = ttk.Frame(parent, style='Card.TFrame')
        self.selected_excel_file = None
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Create the Excel file selection widgets."""
        # Card header
        header_frame = ttk.Frame(self.frame)
        header_frame.pack(fill='x', padx=20, pady=(20, 10))
        
        title_label = ttk.Label(header_frame, text="📊 Excel Output File", style='CardTitle.TLabel')
        title_label.pack(side='left')
        
        # File info and buttons row
        info_button_frame = ttk.Frame(self.frame)
        info_button_frame.pack(fill='x', padx=20, pady=(0, 15))
        
        self.info_label = ttk.Label(info_button_frame, text="Select an Excel file to transfer addresses to", 
                                   style='Body.TLabel')
        self.info_label.pack(side='left')
        
        # Buttons on the right
        button_frame = ttk.Frame(info_button_frame)
        button_frame.pack(side='right')
        
        self.select_btn = ttk.Button(button_frame, text="📂 Choose Excel File",
                                    command=self._select_excel_file,
                                    style='Primary.TButton')
        self.select_btn.pack(side='left', padx=(0, 8))
        
        self.new_btn = ttk.Button(button_frame, text="📄 New File",
                                 command=self._use_default,
                                 style='Secondary.TButton')
        self.new_btn.pack(side='left')
        
        # Description area
        desc_frame = ttk.Frame(self.frame)
        desc_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        desc_text = ("• Choose an existing Excel file to append addresses to\n"
                    "• Or use 'New File' to create a fresh extracted_addresses.xlsx\n"
                    "• Addresses will be added to the first worksheet")
        
        desc_label = ttk.Label(desc_frame, text=desc_text, 
                              style='Body.TLabel',
                              justify='left')
        desc_label.pack(anchor='w')
    
    def _select_excel_file(self):
        """Open file dialog to select an Excel file."""
        file = filedialog.askopenfilename(
            title="Select Excel File to Append Addresses",
            filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls"), ("All files", "*.*")]
        )
        
        if file:
            self.selected_excel_file = file
            filename = os.path.basename(file)
            self.info_label.config(text=f"Will append to: {filename}")
            
            if self.on_selection_change:
                self.on_selection_change(file)
    
    def _use_default(self):
        """Use default new Excel file."""
        self.selected_excel_file = None
        self.info_label.config(text="Will create: extracted_addresses.xlsx")
        
        if self.on_selection_change:
            self.on_selection_change(None)
    
    def get_excel_file(self) -> str:
        """Get the selected Excel file path, or None for default."""
        return self.selected_excel_file
    
    def pack(self, **kwargs):
        """Pack the frame."""
        self.frame.pack(**kwargs)


class MainWindow:
    """Main application window with original design."""
    
    def __init__(self, title: str = "Document Processor Pro"):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("900x700")
        
        # Configure styles
        ModernStyles.configure_styles()
        
        # Set icon
        self._set_icon()
        
        # Configure root
        #self.root.configure(bg='#f8fafc')
        self.root.configure(bg='#ffffff')

        self._create_layout()
    
    def _set_icon(self):
        """Set application icon."""
        try:
            icon_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'assets', 'app_icon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            # Fallback: remove default tkinter icon
            try:
                self.root.wm_iconbitmap("")
            except:
                pass
    
    def _create_layout(self):
        """Create the main layout matching original design."""
        # Main container
        main_container = ttk.Frame(self.root)
        main_container.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Header section with title
        header_frame = ttk.Frame(main_container, style='Card.TFrame')
        header_frame.pack(fill='x', pady=(0, 15))
        
        header_content = ttk.Frame(header_frame)
        header_content.pack(fill='x', padx=25, pady=20)
        
        title_label = ttk.Label(header_content, 
                       text="Document Processor Pro",
                       style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(header_content,
                                  text="Extract addresses from PDFs and print documents",
                                  style='Subtitle.TLabel')
        subtitle_label.pack(pady=(5, 0))
        
        # Main content area
        content_frame = ttk.Frame(main_container)
        content_frame.pack(fill='both', expand=True)
        
        # Left column for file selection
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 8))
        
        # Right column for actions and log
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side='right', fill='both', expand=True, padx=(8, 0))
        
        # Create cards in the original layout
        self.pdf_card = FileSelectionCard(
            left_frame,
            "📁 PDF Folder Selection",
            [("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        self.pdf_card.pack(fill='both', expand=True, pady=(0, 15))
        
        # Excel selection card in left column
        self.excel_card = ExcelSelectionCard(left_frame)
        self.excel_card.pack(fill='x')
        
        # Action card in right column
        self.action_card = ActionCard(right_frame)
        self.action_card.pack(fill='x', pady=(0, 15))
        
        # Log card in right column
        self.log_card = LoggingCard(right_frame)
        self.log_card.pack(fill='both', expand=True)
        
        # Bind clear log button
        self.log_card.clear_btn.configure(command=self.log_card.clear_log)
    
    def run(self):
        """Start the application."""
        self.root.mainloop()
    
    def destroy(self):
        """Destroy the window."""
        self.root.destroy()
