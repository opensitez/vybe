*> vybe-test: cobol/usage_national/national_group_item_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT4.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 NGROUP.
   05 N1 PIC N(5).
   05 N2 PIC N(5).
PROCEDURE DIVISION.
    STOP RUN.

