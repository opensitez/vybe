*> vybe-test: cobol/usage_national/national_initialize_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT6.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(10).
PROCEDURE DIVISION.
    INITIALIZE N1.
    STOP RUN.

