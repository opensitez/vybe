*> vybe-test: cobol/display_formatting/display_decimal_field
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC 9(3)V99 VALUE 123.45.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY D.
    MOVE SPACES TO WS-VYBE-L
    STRING D DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12345"
        DISPLAY "FAIL: want [12345] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

