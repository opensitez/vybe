*> vybe-test: cobol/report_page_handling/report_first_and_last_detail_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH2.
DATA DIVISION.
REPORT SECTION.
RD R1 PAGE LIMIT IS 66 FIRST DETAIL 5 LAST DETAIL 55.
01 D1 TYPE DETAIL.
   03 COL 1 VALUE "X".
PROCEDURE DIVISION.
    STOP RUN.

