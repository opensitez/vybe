*> vybe-test: cobol/nested_if_else/if_four_branches_chained
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N = 1
        DISPLAY "ONE"
    ELSE IF N = 2
        DISPLAY "TWO"
    ELSE IF N = 3
        DISPLAY "THREE"
    ELSE IF N = 4
        DISPLAY "FOUR"
    ELSE
        DISPLAY "OTHER"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "THREE"
        DISPLAY "FAIL: want [THREE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

