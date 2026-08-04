*> vybe-test: cobol/usage_national/national_string_move_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT7.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(10).
PROCEDURE DIVISION.
    MOVE N"WORLD" TO N1.
    STOP RUN.

