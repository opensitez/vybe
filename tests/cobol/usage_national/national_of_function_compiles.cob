*> vybe-test: cobol/usage_national/national_of_function_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT9.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(10).
01 D1 PIC X(10).
PROCEDURE DIVISION.
    MOVE FUNCTION NATIONAL-OF(D1) TO N1.
    STOP RUN.

