*> vybe-test: cobol/math_builtins/add_subtract_and_move_compiles
*> origin: languages/cobol/tests/cobol/test_math_builtins.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 10.
01 WS-B PIC 9(3) VALUE 5.
01 WS-C PIC 9(3).
PROCEDURE DIVISION.
    ADD WS-A TO WS-B.
    SUBTRACT WS-B FROM WS-A.
    MOVE WS-A TO WS-C.
    STOP RUN.

