*> vybe-test: cobol/signed_arithmetic/signed_move_negative_to_signed_field
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC S9(3) VALUE -42.
01 DST PIC S9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    DISPLAY DST.
    MOVE SPACES TO WS-VYBE-L
    STRING DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-00042"
        DISPLAY "FAIL: want [-00042] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

