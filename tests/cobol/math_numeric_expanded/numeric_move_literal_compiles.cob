*> vybe-test: cobol/math_numeric_expanded/numeric_move_literal_compiles
*> origin: languages/cobol/tests/cobol/test_math_numeric_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(5).
PROCEDURE DIVISION.
    MOVE 12345 TO A.
    STOP RUN.

