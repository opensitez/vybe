*> vybe-test: cobol/signed_arithmetic/signed_compute_chain_with_negatives
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC S9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE R = -3 * -4 + -2.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00010"
        DISPLAY "FAIL: want [00010] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

