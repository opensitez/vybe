*> vybe-test: cobol/signed_arithmetic/signed_subtract_negative_from_negative
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(4) VALUE -20.
01 B PIC S9(4) VALUE -8.
01 R PIC S9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SUBTRACT B FROM A GIVING R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0001r"
        DISPLAY "FAIL: want [0001r] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

