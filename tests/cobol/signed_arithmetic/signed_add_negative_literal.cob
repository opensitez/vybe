*> vybe-test: cobol/signed_arithmetic/signed_add_negative_literal
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC S9(4) VALUE +100.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    ADD -35 TO A.
    DISPLAY A.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "+0065"
        DISPLAY "FAIL: want [+0065] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

