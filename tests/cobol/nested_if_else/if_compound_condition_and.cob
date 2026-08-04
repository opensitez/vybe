*> vybe-test: cobol/nested_if_else/if_compound_condition_and
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A = 5 AND B = 5
        DISPLAY "BOTH FIVE"
    ELSE
        DISPLAY "NOT BOTH"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BOTH FIVE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BOTH FIVE"
        DISPLAY "FAIL: want [BOTH FIVE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

