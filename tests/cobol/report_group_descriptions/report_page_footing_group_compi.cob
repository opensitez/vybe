*> vybe-test: cobol/report_group_descriptions/report_page_footing_group_compiles
*> origin: languages/cobol/tests/cobol/test_report_group_descriptions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RGD6.
ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL.
SELECT RPT-FILE ASSIGN TO "rpt.txt".

DATA DIVISION.
FILE SECTION.
FD RPT-FILE REPORT IS SALES-RPT.
REPORT SECTION.
RD SALES-RPT.
01 PAGE-FOOT TYPE PAGE FOOTING.
   03 COL 1 VALUE "FOOT".
PROCEDURE DIVISION.
    STOP RUN.

