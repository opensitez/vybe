*> vybe-test: cobol/report_page_handling/report_page_limit_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH1.
DATA DIVISION.
REPORT SECTION.
RD R1 PAGE LIMIT IS 60.
01 PH TYPE PAGE HEADING.
   03 COL 1 VALUE "PAGE".
PROCEDURE DIVISION.
    STOP RUN.

