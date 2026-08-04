*> vybe-test: cobol/report_controls/report_controls_multiple_control_items_compiles
*> origin: languages/cobol/tests/cobol/test_report_controls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RC5.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE WS-DEPT WS-TEAM.
01 H1 TYPE CONTROL HEADING WS-DEPT.
   03 COL 1 VALUE "D".
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
01 WS-TEAM PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

