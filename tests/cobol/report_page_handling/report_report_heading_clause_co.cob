*> vybe-test: cobol/report_page_handling/report_report_heading_clause_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH8.
DATA DIVISION.
REPORT SECTION.
RD R1.
01 H1 TYPE REPORT HEADING.
   03 COL 1 VALUE "TOP".
PROCEDURE DIVISION.
    STOP RUN.

