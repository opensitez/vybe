*> vybe-test: cobol/report_controls/report_control_footing_compiles
*> origin: languages/cobol/tests/cobol/test_report_controls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RC2.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE WS-DEPT.
01 DEPT-FOOT TYPE CONTROL FOOTING WS-DEPT.
   03 COL 1 VALUE "TOTAL".
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

