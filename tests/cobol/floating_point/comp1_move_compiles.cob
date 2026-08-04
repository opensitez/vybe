*> vybe-test: cobol/floating_point/comp1_move_compiles
*> origin: languages/cobol/tests/cobol/test_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. FP4.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE COMP-1.
01 B USAGE COMP-1.
PROCEDURE DIVISION.
    MOVE A TO B.
    STOP RUN.

