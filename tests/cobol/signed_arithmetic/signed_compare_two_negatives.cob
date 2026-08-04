*> vybe-test: cobol/signed_arithmetic/signed_compare_two_negatives
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(3) VALUE -10.
01 B PIC S9(3) VALUE -5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A < B
        DISPLAY "A LESS"
    ELSE
        DISPLAY "B LESS OR EQUAL"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "A LESS" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A LESS"
        DISPLAY "FAIL: want [A LESS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

