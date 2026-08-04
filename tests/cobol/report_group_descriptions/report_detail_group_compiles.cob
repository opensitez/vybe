*> vybe-test: cobol/report_group_descriptions/report_detail_group_compiles
*> origin: languages/cobol/tests/cobol/test_report_group_descriptions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. RGD2.
DATA DIVISION.
REPORT SECTION.
RD SALES-RPT.
01 DETAIL-LINE TYPE DETAIL.
   03 COL 1 PIC X(10) SOURCE WS-NAME.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(10).
PROCEDURE DIVISION.
    STOP RUN.

