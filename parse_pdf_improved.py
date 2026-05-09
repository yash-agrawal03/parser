import fitz  # PyMuPDF library
import pandas as pd
import re
import os
from datetime import datetime

def parse_student_list_to_excel(pdf_path, excel_path):
    """
    Extracts student selection data from a PDF file and saves it to an Excel file.
    Specifically designed for Maharashtra NEET selection list format.

    Args:
        pdf_path (str): The path to the input PDF file.
        excel_path (str): The path where the output Excel file will be saved.
    """
    # --- 1. Check if PDF file exists ---
    if not os.path.exists(pdf_path):
        print(f"Error: The file '{pdf_path}' was not found.")
        print("Please make sure the PDF file is in the same directory as the script, or provide the full path.")
        return

    print(f"Opening and reading PDF: {pdf_path}...")
    
    # --- 2. Extract text from all pages of the PDF ---
    try:
        doc = fitz.open(pdf_path)
        print(f"PDF has {len(doc)} pages. This may take a few minutes...")
        
        full_text = ""
        processed_pages = 0
        
        for page_num, page in enumerate(doc):
            # Show progress every 100 pages
            if page_num % 100 == 0:
                print(f"Processing page {page_num + 1}/{len(doc)}...")
            
            # Extract text from page
            page_text = page.get_text("text")
            full_text += page_text + "\n"
            processed_pages += 1
            
        doc.close()
        print(f"Successfully processed {processed_pages} pages.")
        
    except Exception as e:
        print(f"An error occurred while reading the PDF file: {e}")
        return

    # --- 3. Define the Regular Expression for this specific format ---
    # Based on the header format:
    # Sr. No. | AIR | NEET Roll No. | CET Form No. | Name | G | Cat. | Quota | Code | College
    # Looking for pattern like: number number number number NAME M/F CATEGORY ...
    
    print("Parsing extracted text to find student data...")
    
    # Split text into lines for easier processing
    lines = full_text.split('\n')
    extracted_data = []
    
    # Find data lines (skip headers and other content)
    data_started = False
    
    for line_num, line in enumerate(lines):
        line = line.strip()
        
        # Skip empty lines
        if not line:
            continue
            
        # Check if we've reached the data section
        if "Sr.    AIR     NEET" in line or "No.            Roll No." in line:
            data_started = True
            continue
            
        # Skip header separator lines
        if line.startswith('---') or 'Legends' in line:
            continue
            
        # Skip if data hasn't started yet
        if not data_started:
            continue
            
        # Try to match student data line
        # Pattern: Sr.No Air NeRollNo CETFormNo Name Gender Category QuotaInfo CollegeCode CollegeName
        # Use a more flexible regex that can handle the specific format
        match = re.match(r'^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([A-Z\s]+?)\s+([MF])\s+(.*)', line)
        
        if match:
            sr_no, air, neet_roll, cet_form, name, gender, rest_of_line = match.groups()
            
            # Clean up the name
            name = name.strip()
            rest_of_line = rest_of_line.strip()
            
            # Default values
            category = "N/A"
            quota = "N/A"
            college_code = "N/A"
            college_name = "N/A"
            
            # Parse the rest of the line to extract category, quota, and college info
            # The format typically is: CATEGORY QUOTA COLLEGECODE CollegeName
            
            if "Choice Not Available" in rest_of_line:
                # Handle case where no college is assigned
                parts = rest_of_line.split("Choice Not Available")
                if parts[0].strip():
                    category_quota = parts[0].strip().split()
                    if category_quota:
                        category = category_quota[0]
                        if len(category_quota) > 1:
                            quota = " ".join(category_quota[1:])
                quota = quota + " Choice Not Available" if quota != "N/A" else "Choice Not Available"
            else:
                # Handle case where college is assigned
                # Look for 4-digit college code followed by colon
                college_match = re.search(r'(\d{4})\s*:\s*(.*?)$', rest_of_line)
                if college_match:
                    college_code = college_match.group(1).strip()
                    college_name = college_match.group(2).strip()
                    
                    # Everything before the college code is category and quota
                    before_college = rest_of_line[:college_match.start()].strip()
                    
                    # Split category and quota
                    cat_quota_parts = before_college.split()
                    if cat_quota_parts:
                        category = cat_quota_parts[0]
                        if len(cat_quota_parts) > 1:
                            quota = " ".join(cat_quota_parts[1:])
                        else:
                            quota = "OPEN"
                else:
                    # If no college code found, treat entire rest as category/quota
                    cat_quota_parts = rest_of_line.split()
                    if cat_quota_parts:
                        category = cat_quota_parts[0]
                        if len(cat_quota_parts) > 1:
                            quota = " ".join(cat_quota_parts[1:])
            
            # Append the structured data
            extracted_data.append({
                "Sr. No.": int(sr_no),
                "AIR": int(air),
                "NEET Roll No.": neet_roll,
                "CET Form No.": cet_form,
                "Name": name,
                "Gender": gender,
                "Category": category,
                "Quota": quota,
                "College Code": college_code,
                "College Name": college_name
            })
            
            # Print progress every 1000 records
            if len(extracted_data) % 1000 == 0:
                print(f"Extracted {len(extracted_data)} records so far...")

    # --- 4. Create a DataFrame and save to Excel ---
    if extracted_data:
        print(f"Found {len(extracted_data)} student records. Creating Excel file...")
        df = pd.DataFrame(extracted_data)
        
        # Ensure long numbers are stored as text in Excel to prevent formatting issues
        df['NEET Roll No.'] = df['NEET Roll No.'].astype(str)
        df['CET Form No.'] = df['CET Form No.'].astype(str)
        df['College Code'] = df['College Code'].astype(str)
        
        # Sort by Sr. No. to maintain order
        df = df.sort_values('Sr. No.')
        
        # Save to Excel with proper formatting
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Student_Selection_List', index=False)
            
            # Get the worksheet to apply formatting
            worksheet = writer.sheets['Student_Selection_List']
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✅ Success! Data has been saved to '{excel_path}'")
        print(f"📊 Total records extracted: {len(extracted_data)}")
        
        # Show some sample data
        print("\n📋 Sample of extracted data:")
        print(df.head().to_string(index=False))
        
    else:
        print("❌ Could not find any matching student data in the PDF.")
        print("This can happen if:")
        print("1. The PDF's text layer is corrupted or has an unexpected format")
        print("2. The data structure is different from expected")
        print("3. The PDF is image-based and needs OCR")

def main():
    """Main function to run the PDF parsing"""
    print("=" * 60)
    print("📄 MAHARASHTRA NEET SELECTION LIST PARSER")
    print("=" * 60)
    print(f"⏰ Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # File names
    pdf_filename = "SellList+R1-MBBS-BDS.pdf"
    excel_filename = "Student_Selection_List_Parsed.xlsx"
    
    # Check if files exist
    if not os.path.exists(pdf_filename):
        print(f"❌ Error: PDF file '{pdf_filename}' not found in current directory.")
        print("📁 Current directory contents:")
        for file in os.listdir('.'):
            print(f"   - {file}")
        return
    
    print(f"📂 Input PDF: {pdf_filename}")
    print(f"📊 Output Excel: {excel_filename}")
    print()
    
    # Run the parsing function
    parse_student_list_to_excel(pdf_filename, excel_filename)
    
    print(f"\n⏰ Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

# --- Main execution block ---
if __name__ == "__main__":
    main()
