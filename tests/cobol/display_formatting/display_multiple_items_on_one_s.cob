*> vybe-test: cobol/display_formatting/display_multiple_items_on_one_statement
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(3) VALUE "FOO".
01 B PIC X(3) VALUE "BAR".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    DISPLAY A " " B.
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE " " DELIMITED SIZE B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FOO BAR"
        DISPLAY "FAIL: want [FOO BAR] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

