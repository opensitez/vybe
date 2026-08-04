*> vybe-test: cobol/report_group_descriptions/report_heading_group_compiles
*> origin: languages/cobol/tests/cobol/test_report_group_descriptions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RGD1.
DATA DIVISION.
REPORT SECTION.
RD SALES-RPT.
01 HEAD-LINE TYPE REPORT HEADING.
   03 COL 1 VALUE "SALES".
PROCEDURE DIVISION.
    STOP RUN.

