*> vybe-test: cobol/nested_if_else/if_compound_condition_or
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 0.
01 B PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A = 5 OR B = 5
        DISPLAY "ONE FIVE"
    ELSE
        DISPLAY "NONE FIVE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE FIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONE FIVE"
        DISPLAY "FAIL: want [ONE FIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

