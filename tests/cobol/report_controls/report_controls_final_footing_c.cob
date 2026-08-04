*> vybe-test: cobol/report_controls/report_controls_final_footing_compiles
*> origin: languages/cobol/tests/cobol/test_report_controls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RC9.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE FINAL.
01 F1 TYPE CONTROL FOOTING FINAL.
   03 COL 1 VALUE "FINAL".
PROCEDURE DIVISION.
    STOP RUN.

