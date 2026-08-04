*> vybe-test: cobol/report_group_descriptions/report_footing_group_compiles
*> origin: languages/cobol/tests/cobol/test_report_group_descriptions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RGD3.
DATA DIVISION.
REPORT SECTION.
RD SALES-RPT.
01 FOOT-LINE TYPE REPORT FOOTING.
   03 COL 1 VALUE "END".
PROCEDURE DIVISION.
    STOP RUN.

