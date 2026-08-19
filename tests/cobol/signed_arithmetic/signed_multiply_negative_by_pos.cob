*> vybe-test: cobol/signed_arithmetic/signed_multiply_negative_by_positive
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(3) VALUE -4.
01 B PIC S9(3) VALUE +7.
01 R PIC S9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MULTIPLY A BY B GIVING R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0002x"
        DISPLAY "FAIL: want [0002x] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

