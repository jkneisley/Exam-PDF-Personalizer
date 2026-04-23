import fitz  # PyMuPDF
import pandas as pd
import os

# --- CONFIGURATION (The "Where" part) ---
# Coordinates are (X, Y). (0,0) is top-left.
FIRST_NAME_POS = (255, 142)   # Box for First Name
LAST_NAME_POS = (255, 177)    # Box for Last Name
ID_POS = (255, 212)           # Box for Student ID
HEADER_POS = (447, 15)        # Odd Pages: Top Right Header
TEST_ID_BOX = (48, 773, 102, 785) # Bottom Left search area

def process_exams(roster_csv, pdf_folder):
    # Load your roster
    df = pd.read_csv(roster_csv)
    updated_rows = []

    for index, student in df.iterrows():
        # Data setup
        first = str(student['First_Name'])
        last = str(student['Last_Name'])
        stu_id = str(student['ID'])
        version = str(student['Version'])
        
        full_name = f"{first} {last}"
        short_header = f"{last}, {first[0]} - {stu_id[-3:]}"
        
        # Open the specific version for this student
        filename = f"{pdf_folder}/Version_{version}.pdf"
        if not os.path.exists(filename):
            print(f"File {filename} not found. Skipping.")
            continue
            
        doc = fitz.open(filename)
        
        # 1. Identify Test ID from bottom left of page 1
        page1 = doc[0]
        test_id = page1.get_text("text", clip=fitz.Rect(TEST_ID_BOX)).strip()
        # Keep only the first 8 characters if there's extra text
        test_id = "".join(filter(str.isalnum, test_id))[:8]

        # 2. Add Test ID to our list to update the roster later
        student['Test_ID'] = test_id
        updated_rows.append(student)

        # 3. Process Pages
        for i, page in enumerate(doc):
            # Page numbers for us start at 0, so Page 1 is index 0 (even)
            # But humans call Page 1, 3, 5 "Odd". 
            # In code: i=0, 2, 4 are the "Odd" pages (1, 3, 5).
            
            if i == 0:
                # Cover Page: Name and ID
                page.insert_text(NAME_POS, full_name, fontsize=12, fontname="helv", color=(0, 0, 0))
                page.insert_text(ID_POS, stu_id, fontsize=12, fontname="helv", color=(0, 0, 0))

            if i % 2 == 0:
                # Top Right Header on all odd pages (1, 3, 5...)
                # First, draw a small white rectangle to prevent overlap
                header_rect = fitz.Rect(475, 15, 600, 40)
                page.draw_rect(header_rect, color=(1, 1, 1), fill=(1, 1, 1))
                page.insert_text(HEADER_POS, short_header, fontsize=10, fontname="helv", color=(0, 0, 0))

        # 4. Save the PDF
        new_filename = f"{last}-{first[0]}-{stu_id}-{version}-{test_id}.pdf"
        doc.save(new_filename)
        doc.close()
        print(f"Created: {new_filename}")

    # Save the updated roster with the found Test IDs
    pd.DataFrame(updated_rows).to_csv("updated_roster.csv", index=False)

# Run the program
process_exams('roster.csv', 'templates')
