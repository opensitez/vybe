*> vybe-test: cobol/report_page_handling/report_page_footing_clause_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH7.
DATA DIVISION.
REPORT SECTION.
RD R1 PAGE LIMIT IS 50.
01 P1 TYPE PAGE FOOTING.
   03 LINE 50 COL 1 VALUE "FOOT".
PROCEDURE DIVISION.
    STOP RUN.

