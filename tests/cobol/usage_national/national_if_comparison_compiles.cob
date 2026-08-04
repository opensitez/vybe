*> vybe-test: cobol/usage_national/national_if_comparison_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT10.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(5).
PROCEDURE DIVISION.
    IF N1 = N"HELLO" DISPLAY "Y" END-IF.
    STOP RUN.

