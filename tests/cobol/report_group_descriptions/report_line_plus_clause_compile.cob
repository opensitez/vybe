*> vybe-test: cobol/report_group_descriptions/report_line_plus_clause_compiles
*> origin: languages/cobol/tests/cobol/test_report_group_descriptions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RGD10.
ENVIRONMENT DIVISION. INPUT-OUTPUT SECTION. FILE-CONTROL.
SELECT RPT-FILE ASSIGN TO "rpt.txt".

DATA DIVISION.
FILE SECTION.
FD RPT-FILE REPORT IS SALES-RPT.
REPORT SECTION.
RD SALES-RPT.
01 D1 TYPE DETAIL.
   03 LINE PLUS 1 COL 1 VALUE "ROW".
PROCEDURE DIVISION.
    STOP RUN.

