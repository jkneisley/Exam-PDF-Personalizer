import fitz  # PyMuPDF
import pandas as pd
import os

# --- CONFIGURATION (The "Where" part) ---
# Coordinates are (X, Y). (0,0) is top-left.
FIRST_NAME_POS = (255, 152)   # Box for First Name
LAST_NAME_POS = (255, 187)    # Box for Last Name
ID_POS = (255, 222)           # Box for Student ID
HEADER_POS = (447, 25)        # Odd Pages: Top Right Header
TEST_ID_BOX = (48, 773, 102, 785) # Bottom Left search area
OUTPUT_FOLDER = "Personalized_Exams"

# Map short course numbers to full course labels used in the red instructor header.
COURSE_LABELS = {
    "14": "MATH 1314 College Algebra",
    "1314": "MATH 1314 College Algebra",
    "16": "MATH 1316 Trigonometry",
    "1316": "MATH 1316 Trigonometry",
    "24": "MATH 1324 Business Math",
    "1324": "MATH 1324 Business Math",
    "50": "MATH 1350 Math for Teachers",
    "1350": "MATH 1350 Math for Teachers",
    "51": "MATH 1351 Math for Teachers",
    "1351": "MATH 1351 Math for Teachers",
}


def get_course_label(course_number):
    course_key = str(course_number).strip()
    return COURSE_LABELS.get(course_key)

def process_exams(roster_csv, big_pdf_path, pages_per_booklet, instructor_info=None, version_filter=None):
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    exam_base_name = os.path.splitext(os.path.basename(big_pdf_path))[0]

    try:
        full_df = pd.read_csv(roster_csv)
        # Filters roster to only students matching the version you typed
        df = full_df[full_df['Version'].astype(str).str.strip() == version_filter].copy()

        if df.empty:
            print(f"!!! Error: No students found in the roster with Version '{version_filter}'.")
            return

        print(f"--- Processing {len(df)} students for Version {version_filter} ---")

    except Exception as e:
        print(f"Error: {e}")
        return

    updated_rows = []
    full_pdf = fitz.open(big_pdf_path)
    total_pages = len(full_pdf)

    for index, student in df.iterrows():
        name_parts = str(student['Student']).split(',', 1)
        last = name_parts[0].strip().upper()
        first = name_parts[1].strip().upper() if len(name_parts) > 1 else ""
        stu_id = str(student['SIS User ID']).strip()
        version_label = str(student['Version']).strip()

        start_page = index * pages_per_booklet
        end_page = start_page + pages_per_booklet

        if end_page > total_pages:
            print(f"!!! Error: PDF ends too soon for {first} {last}.")
            break

        doc = fitz.open()
        doc.insert_pdf(full_pdf, from_page=start_page, to_page=end_page-1)

        # Identify Test ID from first page
        page1 = doc[0]
        test_id_raw = page1.get_text("text", clip=fitz.Rect(TEST_ID_BOX)).strip()
        test_id = "".join(filter(str.isalnum, test_id_raw))[:8]

        student['Test_ID'] = test_id
        updated_rows.append(student)

        short_header = f"{last}, {first[0]} - {stu_id[-3:]}"

        for i, page in enumerate(doc):
            # 1. Page 1 Specific Overlays
            if i == 0:
                # Student Name & ID
                page.insert_text(FIRST_NAME_POS, first, fontsize=14, fontname="helv")
                page.insert_text(LAST_NAME_POS, last, fontsize=14, fontname="helv")
                page.insert_text(ID_POS, stu_id, fontsize=14, fontname="helv")

                # Instructor Header (Red Font - Center Top - PAGE 1 ONLY)
                if instructor_info:
                    instr_rect = fitz.Rect(100, 50, 512, 100)
                    # No white-out box used on Page 1 as per your preference
                    page.insert_textbox(instr_rect, instructor_info,
                                        fontsize=18, fontname="helv",
                                        color=(1, 0, 0), align=1)

            # 2. Student Header (Top Right - Odd Pages 1, 3, 5...)
            # NO WHITE BOX HERE
            if i % 2 == 0:
                page.insert_text(HEADER_POS, short_header, fontsize=9, fontname="helv")

        # Save individual file
        new_filename = f"{last}-{first[0]}-{stu_id}-{version_label}-{test_id}-{exam_base_name}.pdf"
        doc.save(os.path.join(OUTPUT_FOLDER, new_filename))
        doc.close()
        print(f"Created: {new_filename}")

    full_pdf.close()
    pd.DataFrame(updated_rows).to_csv(f"updated_roster_{exam_base_name}.csv", index=False)


def split_pdf_by_pages(big_pdf_path, pages_per_booklet):
    """Split a PDF stack into separate files of a fixed page count, no roster needed."""
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    exam_base_name = os.path.splitext(os.path.basename(big_pdf_path))[0]

    full_pdf = fitz.open(big_pdf_path)
    total_pages = len(full_pdf)

    booklet_num = 1
    start_page = 0
    while start_page < total_pages:
        end_page = min(start_page + pages_per_booklet, total_pages)

        if end_page - start_page < pages_per_booklet:
            print(f"!!! Warning: Final booklet only has {end_page - start_page} page(s) (expected {pages_per_booklet}).")

        doc = fitz.open()
        doc.insert_pdf(full_pdf, from_page=start_page, to_page=end_page - 1)

        new_filename = f"{exam_base_name}-part{booklet_num}.pdf"
        doc.save(os.path.join(OUTPUT_FOLDER, new_filename))
        doc.close()
        print(f"Created: {new_filename}")

        start_page = end_page
        booklet_num += 1

    full_pdf.close()


if __name__ == "__main__":
    print("--- Exam Personalization Tool ---")
    mode = input("1. Mode - (P)ersonalize with roster or just (S)plit by page count? [P/S]: ").strip().lower()

    if mode == 's':
        pdf_in = input("2. PDF stack filename (e.g., Midterm.pdf): ")
        try:
            pg_count = int(input("3. Pages per booklet: "))
            split_pdf_by_pages(pdf_in, pg_count)
        except ValueError:
            print("Invalid page count. Please enter a whole number.")
    else:
        csv_in = input("2. Roster filename (e.g., roster.csv): ")
        pdf_in = input("3. PDF stack filename (e.g., Midterm.pdf): ")

        try:
            pg_count = int(input("4. Pages per booklet: "))
            add_instr = input("5. Add instructor header in red? (yes/no): ").lower()
            instr_text = None
            if add_instr == 'yes':
                course_number = input("   Enter course number (e.g., 14): ")
                course_label = get_course_label(course_number)
                if course_label:
                    instr_text = f"J. Kneisley - {course_label}"
                else:
                    course_name = input("   Enter Course Name: ")
                    instr_text = f"J. Kneisley - MATH {course_number} - {course_name}"

            # --- New Prompt 6 ---
            target_version = input("6. Which version is this PDF stack? (e.g., A, B, or Blue): ").strip()

            # Call the function with all the collected info
            process_exams(csv_in, pdf_in, pg_count, instr_text, target_version)

        except ValueError:
            print("Invalid page count. Please enter a whole number.")