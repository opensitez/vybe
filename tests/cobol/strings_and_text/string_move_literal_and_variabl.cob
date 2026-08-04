*> vybe-test: cobol/strings_and_text/string_move_literal_and_variable_preserves_value
*> origin: languages/cobol/tests/cobol/test_strings_and_text.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(8) VALUE "ALPHA".
01 WS-B PIC X(8) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE WS-A TO WS-B.
    DISPLAY WS-B.
    MOVE SPACES TO WS-VYBE-L
    STRING WS-B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ALPHA"
        DISPLAY "FAIL: want [ALPHA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

