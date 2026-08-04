*> vybe-test: cobol/report_writer/report_writer_generate_detail_compiles
*> origin: languages/cobol/tests/cobol/test_report_writer.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RW10.
DATA DIVISION.
REPORT SECTION.
RD R1.
01 D1 TYPE DETAIL.
   03 COL 1 VALUE "X".
PROCEDURE DIVISION.
    GENERATE D1.
    STOP RUN.

