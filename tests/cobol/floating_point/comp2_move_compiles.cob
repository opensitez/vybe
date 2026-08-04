*> vybe-test: cobol/floating_point/comp2_move_compiles
*> origin: languages/cobol/tests/cobol/test_floating_point.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. FP5.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A USAGE COMP-2.
01 B USAGE COMP-2.
PROCEDURE DIVISION.
    MOVE A TO B.
    STOP RUN.

