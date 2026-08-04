*> vybe-test: cobol/signed_arithmetic/signed_subtract_positive_gives_positive
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(4) VALUE -10.
01 B PIC S9(4) VALUE -30.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    SUBTRACT B FROM A.
    DISPLAY A.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "+0020"
        DISPLAY "FAIL: want [+0020] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

