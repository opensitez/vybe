*> vybe-test: cobol/report_controls/report_sum_clause_compiles
*> origin: languages/cobol/tests/cobol/test_report_controls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RC3.
DATA DIVISION.
REPORT SECTION.
RD R1.
01 D1 TYPE DETAIL.
   03 COL 1 PIC 9(5) SUM WS-AMT.
WORKING-STORAGE SECTION.
01 WS-AMT PIC 9(5).
PROCEDURE DIVISION.
    STOP RUN.

