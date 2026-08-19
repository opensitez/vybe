*> vybe-test: cobol/nested_if_else/nested_if_outer_false_branch
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 0.
01 Y PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF X > 0
        IF Y > 0
            DISPLAY "BOTH POS"
        ELSE
            DISPLAY "X POS Y NOT"
        END-IF
    ELSE
        DISPLAY "X NOT POS"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BOTH POS" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BOTH POS"
        DISPLAY "FAIL: want [BOTH POS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

