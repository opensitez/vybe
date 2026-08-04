*> vybe-test: cobol/report_page_handling/report_page_heading_clause_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH6.
DATA DIVISION.
REPORT SECTION.
RD R1 PAGE LIMIT IS 50.
01 P1 TYPE PAGE HEADING.
   03 LINE 1 COL 1 VALUE "HEAD".
PROCEDURE DIVISION.
    STOP RUN.

