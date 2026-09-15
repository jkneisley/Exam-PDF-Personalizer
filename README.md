# Exam-PDF-Personalizer

0. Run the following code to install the libraries: pip install PyMuPDF pandas

1. Run the following code: python personalize_exams.py

2. At prompt 1, choose a mode:
	- `P` - Personalize with roster (adds student name/ID/header info from a CSV roster)
	- `S` - Split by page count (just splits a PDF stack into equal-sized chunks, no roster needed)

## Mode P: Personalize with roster

1. Upload a roster
	Use a CSV file with the columns Student, SIS User ID, and Version (A/B), one student per row.

2. Upload an exam booklet stack
	Use a stack that contains only one version (for example, all A or all B).

3. Follow the prompts for the roster filename, PDF filename, pages per booklet, instructor header, and version.

Output files are saved to `Personalized_Exams/` and named `LAST-F-ID-VERSION-TESTID-examname.pdf`. An `updated_roster_<examname>.csv` file is also created with the detected Test ID for each student.

## Mode S: Split by page count

1. Upload a PDF stack (any multi-page PDF you want split into equal-sized booklets).

2. Follow the prompts for the PDF filename and pages per booklet.

Output files are saved to `Personalized_Exams/` and named `examname-partN.pdf`. If the total page count isn't evenly divisible by the pages-per-booklet value, the final file will contain the remaining leftover pages.

## Troubleshooting

- Error: Invalid page count. Please enter a whole number.
	Fix: Enter a whole number for pages per booklet (for example, 6).
- Error: No students found in the roster with Version 'X'.
	Fix: Make sure the version prompt matches the exact values in the Version column (for example, A or B).
- Error: PDF ends too soon for FIRST LAST.
	Fix: Confirm total PDF pages = number of students in that version x pages per booklet, and confirm roster/PDF order match.
- Error: [Errno 2] No such file or directory
	Fix: Check the roster and PDF filenames entered at the prompts.
- Warning: Final booklet only has N page(s) (expected M).
	This is expected in Split mode when the total page count isn't evenly divisible by the pages-per-booklet value.

