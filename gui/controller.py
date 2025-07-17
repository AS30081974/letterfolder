"""
Application Controller for Document Processor Pro
Coordinates between GUI, core processing, and business logic.
"""

import os
import threading
import PyPDF2
from typing import List, Callable
from tkinter import messagebox

from gui.components import MainWindow
from core.pdf_processor import PDFProcessor
from core.excel_handler import ExcelHandler
from core.printer import HeadlessPrinter


class AppController:
    """Main application controller that coordinates all components."""
    
    def __init__(self):
        self.window = MainWindow("Document Processor Pro")
        self.pdf_processor = PDFProcessor()
        self.excel_handler = ExcelHandler()
        self.printer = HeadlessPrinter()
        
        # Processing state
        self.is_processing = False
        self.stop_requested = False
        
        self._bind_events()
    
    def _bind_events(self):
        """Bind GUI events to controller methods."""
        # Extract button
        self.window.action_card.extract_btn.configure(
            command=self._extract_addresses
        )
        
        # Print button
        self.window.action_card.print_btn.configure(
            command=self._print_pdfs
        )
        
        # Stop button
        self.window.action_card.stop_btn.configure(
            command=self._stop_processing
        )
        
        # Excel file selection with clear functionality
        self.window.excel_card.on_selection_change = self._handle_excel_selection
    
    def _handle_excel_selection(self, selection):
        """Handle Excel file selection and clear spreadsheet operations."""
        if selection == "CLEAR_SPREADSHEET":
            excel_file = self.window.excel_card.get_excel_file()
            if excel_file:
                try:
                    success = self.excel_handler.clear_spreadsheet(
                        excel_file, 
                        log_callback=self.window.log_card.log
                    )
                    if success:
                        filename = os.path.basename(excel_file)
                        messagebox.showinfo(
                            "Spreadsheet Cleared",
                            f"Successfully cleared: {filename}"
                        )
                    else:
                        messagebox.showerror(
                            "Clear Failed",
                            "Failed to clear the spreadsheet."
                        )
                except Exception as e:
                    messagebox.showerror(
                        "Clear Error",
                        f"Error clearing spreadsheet: {e}"
                    )
    
    def _extract_addresses(self):
        """Extract addresses from selected PDF files."""
        pdf_files = self.window.pdf_card.get_files()
        
        if not pdf_files:
            messagebox.showwarning("No Files", "Please select PDF files first.")
            return
        
        # Start extraction in background thread
        self.stop_requested = False
        thread = threading.Thread(
            target=self._run_extraction,
            args=(pdf_files,),
            daemon=True
        )
        thread.start()
    
    def _print_pdfs(self):
        """Print selected PDF files headlessly."""
        pdf_files = self.window.pdf_card.get_files()
        
        if not pdf_files:
            messagebox.showwarning("No Files", "Please select PDF files first.")
            return
        
        if not self.printer.is_available():
            messagebox.showerror(
                "Printer Unavailable", 
                "Headless printing is not available.\n"
                "Microsoft Edge is required for headless printing."
            )
            return
        
        # Start printing in background thread
        self.stop_requested = False
        thread = threading.Thread(
            target=self._run_printing,
            args=(pdf_files,),
            daemon=True
        )
        thread.start()
    
    def _stop_processing(self):
        """Stop current processing operation."""
        self.stop_requested = True
        self.window.log_card.log("🛑 Stop requested by user...")
    
    def _run_extraction(self, pdf_files: List[str]):
        """Run address extraction in background thread using improved logic from pdf reader."""
        try:
            self._start_processing("🔍 Starting address extraction...")
            
            # Process each PDF file using improved extraction
            pdf_texts = {}  # Dictionary to match pdf reader pattern
            
            for i, pdf_file in enumerate(pdf_files):
                if self.stop_requested:
                    break
                
                filename = os.path.basename(pdf_file)
                self.window.log_card.log(f"📄 [{i+1}/{len(pdf_files)}] Processing {filename}...")
                
                try:
                    # Extract using the improved method similar to pdf reader
                    with open(pdf_file, 'rb') as file:
                        reader = PyPDF2.PdfReader(file)
                        text = ''
                        
                        for page_num in range(len(reader.pages)):
                            page = reader.pages[page_num]
                            text += page.extract_text()
                        
                        # Extract text between ".com" and "Dear" using pdf reader logic
                        start = text.find("uk_team_gbmailgps@lilly.com")
                        end = text.find("Dear")
                        
                        if start != -1 and end != -1:
                            lines = text[start + 5:end].split('\n')
                            lines = [line for line in lines if line != " "]  # Filter like pdf reader
                            extracted_text = '\n'.join(lines[2:])  # Skip first 2 lines like pdf reader
                            
                            if extracted_text.strip():
                                pdf_texts[filename] = extracted_text
                                self.window.log_card.log(f"   ✅ Successfully extracted address")
                            else:
                                self.window.log_card.log(f"   ⚠️ No address content found")
                        else:
                            self.window.log_card.log(f"   ⚠️ Required markers not found")
                        
                except Exception as e:
                    self.window.log_card.log(f"   ❌ Error: {e}")
            
            if not self.stop_requested and pdf_texts:
                # Get the selected Excel file
                excel_file = self.window.excel_card.get_excel_file()
                
                if excel_file:
                    # Use existing Excel file
                    excel_path = excel_file
                    self.window.log_card.log(f"\n💾 Saving {len(pdf_texts)} addresses to existing Excel file...")
                else:
                    # Use default filename like pdf reader
                    excel_path = "addresses.xlsx"
                    self.window.log_card.log(f"\n💾 Saving {len(pdf_texts)} addresses to {excel_path}...")
                
                # Save using the Excel handler method that matches pdf reader approach
                success = self.excel_handler.save_data_to_excel(
                    pdf_texts, 
                    excel_path,
                    log_callback=self.window.log_card.log
                )
                
                if success:
                    filename = os.path.basename(excel_path)
                    self.window.log_card.log(f"✅ Addresses saved to {filename}")
                    messagebox.showinfo(
                        "Extraction Complete",
                        f"Successfully extracted {len(pdf_texts)} addresses.\n"
                        f"Saved to: {filename}"
                    )
                else:
                    self.window.log_card.log("❌ Failed to save addresses")
                    messagebox.showerror("Save Error", "Failed to save addresses to Excel.")
            
            elif not self.stop_requested:
                self.window.log_card.log("⚠️ No addresses found in any files")
                messagebox.showinfo("No Results", "No addresses were found in the selected files.")
        
        except Exception as e:
            self.window.log_card.log(f"❌ Extraction failed: {e}")
            messagebox.showerror("Extraction Error", f"Extraction failed: {e}")
        
        finally:
            self._finish_processing()
    
    def _run_printing(self, pdf_files: List[str]):
        """Run PDF printing in background thread."""
        try:
            self._start_processing("🖨️ Starting headless printing...")
            
            successful, failed = self.printer.print_multiple_pdfs(
                pdf_files,
                log_callback=self.window.log_card.log,
                stop_callback=lambda: self.stop_requested
            )
            
            if not self.stop_requested:
                self.window.log_card.log(f"\n📊 Printing complete:")
                self.window.log_card.log(f"   ✅ Successful: {successful}")
                self.window.log_card.log(f"   ❌ Failed: {failed}")
                
                if successful > 0:
                    messagebox.showinfo(
                        "Printing Complete",
                        f"Successfully printed {successful} files.\n"
                        f"Failed: {failed}"
                    )
                else:
                    messagebox.showerror(
                        "Printing Failed",
                        "No files were printed successfully."
                    )
        
        except Exception as e:
            self.window.log_card.log(f"❌ Printing failed: {e}")
            messagebox.showerror("Printing Error", f"Printing failed: {e}")
        
        finally:
            self._finish_processing()
    
    def _start_processing(self, message: str):
        """Start processing mode."""
        self.is_processing = True
        
        # Update GUI state
        self.window.action_card.extract_btn.configure(state='disabled')
        self.window.action_card.print_btn.configure(state='disabled')
        self.window.action_card.stop_btn.configure(state='normal')
        
        # Show progress and log
        self.window.log_card.show_progress()
        self.window.log_card.log(message)
    
    def _finish_processing(self):
        """Finish processing mode."""
        self.is_processing = False
        self.stop_requested = False
        
        # Update GUI state
        self.window.action_card.extract_btn.configure(state='normal')
        self.window.action_card.print_btn.configure(state='normal')
        self.window.action_card.stop_btn.configure(state='disabled')
        
        # Hide progress
        self.window.log_card.hide_progress()
        self.window.log_card.log("✅ Operation complete.\n")
    
    def run(self):
        """Start the application."""
        self.window.run()
    
    def stop(self):
        """Stop the application."""
        if self.is_processing:
            self.stop_requested = True
        self.window.destroy()
