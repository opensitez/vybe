*> vybe-test: cobol/report_page_handling/report_report_footing_clause_compiles
*> origin: languages/cobol/tests/cobol/test_report_page_handling.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RPH9.
ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL.
SELECT RPT-FILE ASSIGN TO "rpt.txt".

DATA DIVISION.
FILE SECTION.
FD RPT-FILE REPORT IS R1.
REPORT SECTION.
RD R1.
01 F1 TYPE REPORT FOOTING.
   03 COL 1 VALUE "END".
PROCEDURE DIVISION.
    STOP RUN.

