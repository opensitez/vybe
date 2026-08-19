*> vybe-test: cobol/compute_expressions/compute_multiply_two_vars
*> origin: languages/cobol/tests/cobol/test_compute_expressions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9(3) VALUE 12.
01 B PIC 9(3) VALUE 8.
01 R PIC 9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE R = A * B.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00096"
        DISPLAY "FAIL: want [00096] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

