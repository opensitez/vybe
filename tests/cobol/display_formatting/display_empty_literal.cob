*> vybe-test: cobol/display_formatting/display_empty_literal
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY "".
    MOVE SPACES TO WS-VYBE-L
    STRING "" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = ""
        DISPLAY "FAIL: want [] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

