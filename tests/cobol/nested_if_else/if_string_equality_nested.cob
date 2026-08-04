*> vybe-test: cobol/nested_if_else/if_string_equality_nested
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 CODE PIC X(2) VALUE "OK".
01 TYPE PIC X VALUE "A".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF CODE = "OK"
        IF TYPE = "A"
            DISPLAY "OK-A"
        ELSE
            DISPLAY "OK-OTHER"
        END-IF
    ELSE
        DISPLAY "NOT-OK"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "OK-A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OK-A"
        DISPLAY "FAIL: want [OK-A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

