*> vybe-test: cobol/report_controls/report_controls_detail_with_source_compiles
*> origin: languages/cobol/tests/cobol/test_report_controls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RC8.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE WS-DEPT.
01 D1 TYPE DETAIL.
   03 COL 1 SOURCE WS-NAME.
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
01 WS-NAME PIC X(10).
PROCEDURE DIVISION.
    STOP RUN.

