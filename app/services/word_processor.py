import win32com.client as win32
import pythoncom
import os
from typing import Optional

class WordProcessor:
    def __init__(self):
        self.word = None
        self.doc = None

    def __enter__(self):
        pythoncom.CoInitialize()
        try:
            self.word = win32.Dispatch('Word.Application')
            self.word.Visible = False
            return self
        except Exception as e:
            pythoncom.CoUninitialize()
            raise e

    def __exit__(self, exc_type, exc_val, exc_tb):
        try:
            if self.doc:
                self.doc.Close()
            if self.word:
                self.word.Quit()
        finally:
            pythoncom.CoUninitialize()

    def process_document(self, input_path: str, output_path: str, hidden_text: str, 
                        font_size: int = 2) -> Optional[str]:
        try:
            input_path = os.path.abspath(input_path)
            output_path = os.path.abspath(output_path)
            
            if not os.path.exists(input_path):
                raise FileNotFoundError(f"Input file not found: {input_path}")
            
            self.doc = self.word.Documents.Open(input_path)
            
            # Get page dimensions
            page_width = self.doc.PageSetup.PageWidth
            page_height = self.doc.PageSetup.PageHeight
            
            # Add textbox with exact positioning
            shape = self.doc.Shapes.AddTextbox(
                1,  # Orientation
                page_width - 200,  # Left position
                page_height - 40,  # Top position
                200,  # Width
                40   # Height
            )
            
            # Configure textbox exactly as needed
            shape.TextFrame.TextRange.Text = hidden_text
            shape.TextFrame.TextRange.Font.Size = font_size
            shape.Fill.Visible = False
            shape.Line.Visible = False
            shape.TextFrame.TextRange.Font.Color = 16777215  # White color
            shape.WrapFormat.Type = 3  # Behind text
            
            self.doc.SaveAs(output_path)
            return output_path
            
        except Exception as e:
            print(f"Error processing document: {str(e)}")
            return None 