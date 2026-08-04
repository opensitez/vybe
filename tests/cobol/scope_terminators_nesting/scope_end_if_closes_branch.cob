*> vybe-test: cobol/scope_terminators_nesting/scope_end_if_closes_branch
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N > 0
        DISPLAY "POS"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "POS" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "POS"
        DISPLAY "FAIL: want [POS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY "AFTER".
    MOVE SPACES TO WS-VYBE-L
    STRING "AFTER" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AFTER"
        DISPLAY "FAIL: want [AFTER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

