*> vybe-test: cobol/report_group_descriptions/report_page_heading_group_compiles
*> origin: languages/cobol/tests/cobol/test_report_group_descriptions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RGD5.
DATA DIVISION.
REPORT SECTION.
RD SALES-RPT.
01 PAGE-HEAD TYPE PAGE HEADING.
   03 COL 1 VALUE "HEADER".
PROCEDURE DIVISION.
    STOP RUN.

