"""
PDF Processing Module
Handles PDF text extraction and data processing.
"""

import PyPDF2
import glob
import os
from typing import Dict, List


class PDFProcessor:
    """Handles PDF text extraction and processing."""
    
    def __init__(self):
        self.pdf_texts = {}
        
    def extract_text_from_pdfs(self, folder_path: str, log_callback=None) -> Dict[str, str]:
        """
        Extract text from PDF files in the specified folder.
        
        Args:
            folder_path: Path to folder containing PDF files
            log_callback: Optional callback function for logging messages
            
        Returns:
            Dictionary mapping PDF filenames to extracted text
        """
        pdf_texts = {}
        
        # Find all PDF files in the folder
        pdf_pattern = os.path.join(folder_path, "*.pdf")
        pdf_files = glob.glob(pdf_pattern)
        
        if not pdf_files:
            if log_callback:
                log_callback("No PDF files found in the selected folder.")
            return pdf_texts
        
        if log_callback:
            log_callback(f"Found {len(pdf_files)} PDF files to process.")
        
        for file_path in pdf_files:
            try:
                filename = os.path.basename(file_path)
                if log_callback:
                    log_callback(f"Processing: {filename}")
                
                with open(file_path, 'rb') as pdf_file:
                    reader = PyPDF2.PdfReader(pdf_file)
                    text = ''
                    
                    for page_num in range(len(reader.pages)):
                        page = reader.pages[page_num]
                        text += page.extract_text()
                    
                    # Extract text between ".com" and "Dear"
                    extracted_text = self._extract_address_section(text)
                    
                    if extracted_text:
                        pdf_texts[filename] = extracted_text
                        if log_callback:
                            log_callback(f"Successfully extracted data from {filename}")
                    else:
                        if log_callback:
                            log_callback(f"Could not find data markers in {filename}")
                        
            except Exception as e:
                if log_callback:
                    log_callback(f"Error processing {os.path.basename(file_path)}: {str(e)}")
        
        return pdf_texts
    
    def _extract_address_section(self, text: str) -> str:
        """
        Extract the address section from PDF text using improved logic from pdf reader.
        
        Args:
            text: Full PDF text content
            
        Returns:
            Extracted address text or empty string if not found
        """
        start = text.find("uk_team_gbmailgps@lilly.com")
        end = text.find("Dear")
        
        if start != -1 and end != -1:
            # Extract text between the markers
            lines = text[start + 5:end].split('\n')
            # Filter out empty lines and whitespace-only lines
            lines = [line.strip() for line in lines if line.strip() != ""]
            # Skip the first 2 lines as per pdf reader logic, then join the rest
            extracted_text = '\n'.join(lines[2:]) if len(lines) > 2 else '\n'.join(lines)
            return extracted_text
        
        return ""
    
    def extract_addresses(self, pdf_file: str) -> List[str]:
        """
        Extract addresses from a single PDF file.
        
        Args:
            pdf_file: Path to the PDF file
            
        Returns:
            List of extracted address strings
        """
        try:
            with open(pdf_file, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ''
                
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    text += page.extract_text()
                
                # Extract text between ".com" and "Dear" using improved logic
                extracted_text = self._extract_address_section(text)
                
                if extracted_text:
                    return [extracted_text]  # Return as list for consistency
                else:
                    return []
                    
        except Exception as e:
            raise Exception(f"Error processing PDF {pdf_file}: {str(e)}")

    def get_pdf_files(self, folder_path: str) -> List[str]:
        """
        Get list of PDF files in the specified folder.
        
        Args:
            folder_path: Path to folder to search
            
        Returns:
            List of PDF file paths
        """
        pdf_pattern = os.path.join(folder_path, "*.pdf")
        return glob.glob(pdf_pattern)
