*> vybe-test: cobol/report_writer/report_writer_page_limit_compiles
*> origin: languages/cobol/tests/cobol/test_report_writer.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RW9.
DATA DIVISION.
REPORT SECTION.
RD R1 PAGE LIMIT IS 60.
01 D1 TYPE DETAIL.
   03 COL 1 VALUE "X".
PROCEDURE DIVISION.
    STOP RUN.

