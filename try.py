import fitz  # PyMuPDF library
import pandas as pd
import re
import os
from datetime import datetime


def parse_student_list_to_excel(pdf_path, excel_path):
    """
    Extracts student selection data from a PDF file and saves it to an Excel file.
    Specifically designed for Maharashtra NEET selection list format.

    Fixes added:
    ✅ Names starting with '$'
    ✅ Multiline wrapped entries
    ✅ Long college names extending to next line
    ✅ Wrapped quota/category text
    """

    # --- 1. Check if PDF file exists ---
    if not os.path.exists(pdf_path):
        print(f"Error: The file '{pdf_path}' was not found.")
        return

    print(f"Opening and reading PDF: {pdf_path}...")

    # --- 2. Extract text from PDF ---
    try:
        doc = fitz.open(pdf_path)

        print(f"PDF has {len(doc)} pages. This may take a few minutes...")

        full_text = ""

        for page_num, page in enumerate(doc):

            if page_num % 100 == 0:
                print(f"Processing page {page_num + 1}/{len(doc)}...")

            page_text = page.get_text("text")
            full_text += page_text + "\n"

        doc.close()

    except Exception as e:
        print(f"Error reading PDF: {e}")
        return

    print("Parsing extracted text...")

    # --- 3. Split into raw lines ---
    raw_lines = full_text.split('\n')

    # --- 4. Merge wrapped/multiline entries ---
    merged_lines = []
    current_line = ""
    data_started = False

    for line in raw_lines:

        line = line.strip()

        if not line:
            continue

        # Detect header start
        if "Sr." in line and "AIR" in line and "NEET" in line:
            data_started = True
            continue

        if not data_started:
            continue

        # Skip separators / legends
        if line.startswith('---') or 'Legends' in line:
            continue

        # New entry starts with:
        # SrNo AIR NEETROLL CETFORM
        if re.match(r'^\d+\s+\d+\s+\d+\s+\d+', line):

            # Save previous accumulated record
            if current_line:
                merged_lines.append(current_line.strip())

            current_line = line

        else:
            # Continuation of previous line
            current_line += " " + line

    # Add last record
    if current_line:
        merged_lines.append(current_line.strip())

    print(f"Total merged records found: {len(merged_lines)}")

    extracted_data = []

    # Known categories
    known_categories = [
        "OPEN", "OBC", "SC", "ST", "EWS",
        "NT1", "NT2", "NT3",
        "VJ", "VJA",
        "NTB", "NTC", "NTD",
        "SBC", "MIN",
        "DEF1", "DEF2", "DEF3",
        "PWD", "EBC"
    ]

    def extract_category_quota(text):

        text = text.strip()

        for cat in known_categories:

            if text == cat:
                return cat, ""

            elif text.startswith(cat + " "):
                remaining = text[len(cat):].strip()
                remaining = re.sub(r'\s+', ' ', remaining)
                return cat, remaining

            elif text.startswith(cat + "("):
                remaining = text[len(cat):].strip()
                return cat, remaining

        # fallback
        parts = text.split(maxsplit=1)

        if len(parts) == 2:
            return parts[0], re.sub(r'\s+', ' ', parts[1])

        elif len(parts) == 1:
            return parts[0], ""

        return "N/A", ""

    def cleanup_quota(quota_str):

        if not quota_str or quota_str == "N/A":
            return quota_str

        # Fix missing brackets
        quota_str = re.sub(r'\(([A-Za-z])\s*$', r'(\1)', quota_str)
        quota_str = re.sub(r'\(([A-Za-z]{2,})\s*$', r'(\1)', quota_str)

        # Normalize spaces
        quota_str = re.sub(r'\s+', ' ', quota_str).strip()

        return quota_str

    # --- 5. Parse merged lines ---
    for line_num, line in enumerate(merged_lines):

        # Regex FIXED:
        # Allows:
        # $NAME
        # dots
        # apostrophes
        # hyphens
        match = re.match(
            r'^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+([\$A-Z\s\.\'-]+?)\s+([MF])\s+(.*)',
            line
        )

        if not match:
            continue

        sr_no, air, neet_roll, cet_form, name, gender, rest_of_line = match.groups()

        name = re.sub(r'\s+', ' ', name.strip())
        rest_of_line = re.sub(r'\s+', ' ', rest_of_line.strip())

        category = "N/A"
        quota = "N/A"
        college_code = "N/A"
        college_name = "N/A"

        # Case: Choice Not Available
        if "Choice Not Available" in rest_of_line:

            parts = rest_of_line.split("Choice Not Available")

            cat, quota_part = extract_category_quota(parts[0].strip())

            category = cat

            quota_part = quota_part.strip()

            if quota_part:
                quota = quota_part + " Choice Not Available"
            else:
                quota = "Choice Not Available"

            quota = cleanup_quota(quota)

        else:

            # College pattern
            college_match = re.search(
                r'(\d{4})\s*:\s*(.+)$',
                rest_of_line
            )

            if college_match:

                college_code = college_match.group(1).strip()
                college_name = college_match.group(2).strip()

                before_college = rest_of_line[:college_match.start()].strip()

                cat, quota_part = extract_category_quota(before_college)

                category = cat

                quota_part = quota_part.strip()

                quota = quota_part if quota_part else "OPEN"

                quota = cleanup_quota(quota)

            else:

                # No college found
                cat, quota_part = extract_category_quota(rest_of_line)

                category = cat

                quota = quota_part if quota_part else "OPEN"

                quota = cleanup_quota(quota)

        extracted_data.append({
            "Sr. No.": int(sr_no),
            "AIR": int(air),
            "NEET Roll No.": str(neet_roll),
            "CET Form No.": str(cet_form),
            "Name": name,
            "Gender": gender,
            "Category": category,
            "Quota": quota,
            "College Code": str(college_code),
            "College Name": college_name
        })

        if len(extracted_data) % 1000 == 0:
            print(f"Extracted {len(extracted_data)} records...")

    # --- 6. Save to Excel ---
    if extracted_data:

        print(f"Found {len(extracted_data)} records.")

        df = pd.DataFrame(extracted_data)

        df = df.sort_values('Sr. No.')

        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:

            df.to_excel(
                writer,
                sheet_name='Student_Selection_List',
                index=False
            )

            worksheet = writer.sheets['Student_Selection_List']

            # Auto width
            for column in worksheet.columns:

                max_length = 0
                column_letter = column[0].column_letter

                for cell in column:
                    try:
                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )
                    except:
                        pass

                worksheet.column_dimensions[column_letter].width = min(
                    max_length + 2,
                    60
                )

        print(f"\n✅ Excel saved successfully: {excel_path}")
        print(f"📊 Total extracted records: {len(df)}")

        print("\n📋 Sample Data:")
        print(df.head().to_string(index=False))

    else:
        print("❌ No records extracted.")


def main():

    print("=" * 60)
    print("📄 MAHARASHTRA NEET SELECTION LIST PARSER")
    print("=" * 60)

    print(f"⏰ Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Your exact PDF file
    pdf_filename = r"4-NEET UG 2025 MAHARASHTRA AYUSH COURSES 4TH ROUND SELECTION LIST.pdf"

    # Output Excel filename
    excel_filename = "Parsed_Output.xlsx"

    print(f"📂 Input PDF: {pdf_filename}")
    print(f"📊 Output Excel: {excel_filename}")
    print()

    parse_student_list_to_excel(
        pdf_filename,
        excel_filename
    )

    print(f"\n⏰ Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)


# --- Main ---
if __name__ == "__main__":
    main()