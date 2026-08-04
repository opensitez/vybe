*> vybe-test: cobol/report_page_handling/report_line_number_clause_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH4.
DATA DIVISION.
REPORT SECTION.
RD R1.
01 D1 TYPE DETAIL.
   03 LINE NUMBER 3 COL 1 VALUE "ROW".
PROCEDURE DIVISION.
    STOP RUN.

