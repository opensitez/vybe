*> vybe-test: cobol/report_writer/report_writer_control_heading_compiles
*> origin: languages/cobol/tests/cobol/test_report_writer.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RW5.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE WS-DEPT.
01 C1 TYPE CONTROL HEADING WS-DEPT.
   03 COL 1 VALUE "C".
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

