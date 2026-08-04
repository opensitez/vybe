*> vybe-test: cobol/report_controls/report_control_heading_compiles
*> origin: languages/cobol/tests/cobol/test_report_controls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RC1.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE WS-DEPT.
01 DEPT-HEAD TYPE CONTROL HEADING WS-DEPT.
   03 COL 1 VALUE "DEPT".
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

