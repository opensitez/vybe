*> vybe-test: cobol/report_group_descriptions/report_control_heading_group_compiles
*> origin: languages/cobol/tests/cobol/test_report_group_descriptions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RGD7.
DATA DIVISION.
REPORT SECTION.
RD SALES-RPT CONTROLS ARE WS-DEPT.
01 CH1 TYPE CONTROL HEADING WS-DEPT.
   03 COL 1 VALUE "CH".
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

