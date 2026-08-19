*> vybe-test: cobol/signed_arithmetic/signed_v99_positive_and_negative_display
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9V99 VALUE +3.50.
01 B PIC S9V99 VALUE -1.25.
01 R PIC S9(2)V99 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    ADD A B GIVING R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0225"
        DISPLAY "FAIL: want [0225] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

