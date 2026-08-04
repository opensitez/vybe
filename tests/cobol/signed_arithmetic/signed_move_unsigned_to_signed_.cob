*> vybe-test: cobol/signed_arithmetic/signed_move_unsigned_to_signed_field
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC PIC 9(3) VALUE 123.
01 DST PIC S9(4) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    DISPLAY DST.
    MOVE SPACES TO WS-VYBE-L
    STRING DST DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "+0123"
        DISPLAY "FAIL: want [+0123] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

