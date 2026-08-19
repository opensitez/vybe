*> vybe-test: cobol/signed_arithmetic/signed_compare_negative_to_zero
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC S9(3) VALUE -1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N = 0
        DISPLAY "ZERO"
    ELSE
        DISPLAY "NOT ZERO"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ZERO" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZERO"
        DISPLAY "FAIL: want [ZERO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

