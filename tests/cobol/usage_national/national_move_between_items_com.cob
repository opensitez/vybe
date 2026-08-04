*> vybe-test: cobol/usage_national/national_move_between_items_compiles
*> origin: languages/cobol/tests/cobol/test_usage_national.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. NAT5.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N1 PIC N(5).
01 N2 PIC N(5).
PROCEDURE DIVISION.
    MOVE N1 TO N2.
    STOP RUN.

