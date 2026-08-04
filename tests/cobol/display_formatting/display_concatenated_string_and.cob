*> vybe-test: cobol/display_formatting/display_concatenated_string_and_number
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(3) VALUE 42.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY "VALUE: " N.
    MOVE SPACES TO WS-VYBE-L
    STRING "VALUE: " DELIMITED SIZE N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "VALUE: 042"
        DISPLAY "FAIL: want [VALUE: 042] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

