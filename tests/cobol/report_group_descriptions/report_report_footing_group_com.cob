*> vybe-test: cobol/report_group_descriptions/report_report_footing_group_compiles
*> origin: languages/cobol/tests/cobol/test_report_group_descriptions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RGD9.
DATA DIVISION.
REPORT SECTION.
RD SALES-RPT.
01 RF1 TYPE REPORT FOOTING.
   03 COL 1 VALUE "DONE".
PROCEDURE DIVISION.
    STOP RUN.

