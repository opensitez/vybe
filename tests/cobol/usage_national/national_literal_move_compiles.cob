*> vybe-test: cobol/usage_national/national_literal_move_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT2.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(10).
PROCEDURE DIVISION.
    MOVE N"HELLO" TO N1.
    STOP RUN.

