*> vybe-test: cobol/report_controls/report_reset_on_compiles
*> origin: languages/cobol/tests/cobol/test_report_controls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RC4.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE WS-DEPT.
01 D1 TYPE DETAIL.
   03 COL 1 PIC 9(5) SUM WS-AMT RESET ON WS-DEPT.
WORKING-STORAGE SECTION.
01 WS-AMT PIC 9(5).
01 WS-DEPT PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

