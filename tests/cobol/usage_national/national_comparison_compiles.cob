*> vybe-test: cobol/usage_national/national_comparison_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT3.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(5).
PROCEDURE DIVISION.
    IF N1 = N"A" DISPLAY "Y" END-IF.
    STOP RUN.

