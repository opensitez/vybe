*> vybe-test: cobol/report_page_handling/report_page_limit_and_controls_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH10.
DATA DIVISION.
REPORT SECTION.
RD R1 PAGE LIMIT IS 60 CONTROLS ARE WS-DEPT.
01 D1 TYPE DETAIL.
   03 COL 1 VALUE "X".
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

