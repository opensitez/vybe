*> vybe-test: cobol/signed_arithmetic/signed_condition_greater_than_zero
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC S9(3) VALUE +5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N > 0
        DISPLAY "POS"
    ELSE
        DISPLAY "NEG"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "POS" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "POS"
        DISPLAY "FAIL: want [POS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

