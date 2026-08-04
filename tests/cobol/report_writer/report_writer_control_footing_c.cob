*> vybe-test: cobol/report_writer/report_writer_control_footing_compiles
*> origin: languages/cobol/tests/cobol/test_report_writer.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RW6.
DATA DIVISION.
REPORT SECTION.
RD R1 CONTROLS ARE WS-DEPT.
01 C1 TYPE CONTROL FOOTING WS-DEPT.
   03 COL 1 VALUE "T".
WORKING-STORAGE SECTION.
01 WS-DEPT PIC X(4).
PROCEDURE DIVISION.
    STOP RUN.

