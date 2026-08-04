*> vybe-test: cobol/usage_national/usage_national_definition_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT1.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 USAGE NATIONAL PIC N(10).
PROCEDURE DIVISION.
    STOP RUN.

