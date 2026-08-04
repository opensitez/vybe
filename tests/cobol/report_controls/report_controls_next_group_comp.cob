*> vybe-test: cobol/report_controls/report_controls_next_group_compiles
*> origin: languages/cobol/tests/cobol/test_report_controls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RC10.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE WS-DEPT.
01 D1 TYPE DETAIL.
   03 NEXT GROUP PLUS 1.
   03 COL 1 VALUE "ROW".
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

