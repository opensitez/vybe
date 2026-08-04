*> vybe-test: cobol/report_writer/report_writer_page_footing_compiles
*> origin: languages/cobol/tests/cobol/test_report_writer.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RW4.
DATA DIVISION.
REPORT SECTION.
RD R1.
01 P1 TYPE PAGE FOOTING.
   03 COL 1 VALUE "FOOT".
PROCEDURE DIVISION.
    STOP RUN.

