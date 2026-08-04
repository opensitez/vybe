*> vybe-test: cobol/report_writer/report_writer_report_footing_compiles
*> origin: languages/cobol/tests/cobol/test_report_writer.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RW7.
DATA DIVISION.
REPORT SECTION.
RD R1.
01 F1 TYPE REPORT FOOTING.
   03 COL 1 VALUE "END".
PROCEDURE DIVISION.
    STOP RUN.

