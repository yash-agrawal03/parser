import fitz  # PyMuPDF library
import pandas as pd
import re
import os

def parse_student_list_to_excel(pdf_path, excel_path):
    """
    Extracts student selection data from a PDF file and saves it to an Excel file.
    This revised version uses a more robust regex to handle spacing and column variations
    based on the visual structure of the PDF.

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
        full_text = ""
        for page_num, page in enumerate(doc):
            # Using get_text("text") is generally reliable for table-like structures
            full_text += page.get_text("text")
        doc.close()
    except Exception as e:
        print(f"An error occurred while reading the PDF file: {e}")
        return

    # --- 3. Parse text line by line for better control ---
    print("Parsing extracted text to find student data...")
    
    lines = full_text.split('\n')
    extracted_data = []
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
        match = re.match(r'^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([A-Z\s]+?)\s+([MF])\s+(.*)', line)
        
        if match:
            # Show progress every 1000 records
            if len(extracted_data) % 1000 == 0 and len(extracted_data) > 0:
                print(f"Extracted {len(extracted_data)} records so far...")
        else:
            continue
        # Unpack the initial groups captured by the regex
        sr_no, air, neet_roll, cet_form, name, gender, rest_of_line = match.groups()

        # Clean up the extracted fields
        name = name.strip()
        rest_of_line = rest_of_line.strip()

        # Default values for the remaining columns
        category = "N/A"
        quota = "N/A"
        college_code = "N/A"
        college_name = "N/A"

        # Now, intelligently parse the 'rest_of_line'
        if "Choice Not Available" in rest_of_line:
            # Handle the simple case where no college is assigned
            category_parts = rest_of_line.split("Choice Not Available")
            category = category_parts[0].strip()
            quota = "Choice Not Available"
        else:
            # Handle the more complex case where a college is assigned
            # The category is the first part of the string. It can be one or more words.
            # The college info is always at the end, identified by a 4-digit code.
            college_match = re.search(r'(\d{4})\s*:\s*(.*)', rest_of_line)
            if college_match:
                college_code = college_match.group(1).strip()
                college_name = college_match.group(2).strip()
                
                # Everything before the college match is category and quota
                before_college = rest_of_line[:college_match.start()].strip()
                
                # The first word(s) are the category. The rest is the quota.
                # This is an approximation that should work for this format.
                parts = before_college.split(maxsplit=1)
                category = parts[0] if parts else ""
                quota = parts[1] if len(parts) > 1 else "OPEN" # Default to OPEN if quota is not explicitly listed
            else:
                # Fallback if college pattern isn't found
                category = rest_of_line

        # Append the fully structured data to our list
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

    # --- 5. Create a DataFrame and save to Excel ---
    if extracted_data:
        print(f"Found {len(extracted_data)} student records. Creating Excel file...")
        df = pd.DataFrame(extracted_data)
        
        # Ensure long numbers are stored as text in Excel to prevent formatting issues
        df['NEET Roll No.'] = df['NEET Roll No.'].astype(str)
        df['CET Form No.'] = df['CET Form No.'].astype(str)

        df.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"✅ Success! Data has been saved to '{excel_path}'")
    else:
        print("Could not find any matching student data in the PDF.")
        print("This can happen if the PDF's text layer is corrupted or has an unexpected format.")

# --- Main execution block ---
if __name__ == "__main__":
    # Name of your input PDF file.
    pdf_filename = "SellList+R1-MBBS-BDS.pdf"

    # Desired name for the output Excel file.
    excel_filename = "Student_Selection_List_v2.xlsx"

    # Run the parsing function
    parse_student_list_to_excel(pdf_filename, excel_filename)
