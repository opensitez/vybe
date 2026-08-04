*> vybe-test: cobol/report_page_handling/report_page_heading_and_footing_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH3.
DATA DIVISION.
REPORT SECTION.
RD R1.
01 PH TYPE PAGE HEADING.
   03 COL 1 VALUE "H".
01 PF TYPE PAGE FOOTING.
   03 COL 1 VALUE "F".
PROCEDURE DIVISION.
    STOP RUN.

