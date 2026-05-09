import fitz  # PyMuPDF
import os

def test_pdf_extraction(pdf_path):
    """Test PDF text extraction and show first few lines"""
    if not os.path.exists(pdf_path):
        print(f"Error: PDF file '{pdf_path}' not found.")
        return
    
    try:
        doc = fitz.open(pdf_path)
        print(f"PDF has {len(doc)} pages")
        
        # Get text from first page
        page = doc[0]
        text = page.get_text("text")
        
        # Show first 1000 characters
        print("First 1000 characters of PDF text:")
        print("=" * 50)
        print(text[:1000])
        print("=" * 50)
        
        # Show lines to understand structure
        lines = text.split('\n')
        print(f"\nFirst 20 lines:")
        for i, line in enumerate(lines[:20]):
            print(f"{i+1:2d}: {repr(line)}")
        
        doc.close()
        
    except Exception as e:
        print(f"Error reading PDF: {e}")

if __name__ == "__main__":
    test_pdf_extraction("SellList+R1-MBBS-BDS.pdf")
