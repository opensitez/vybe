*> vybe-test: cobol/scope_terminators_nesting/scope_end_if_else_both_branches
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF F = 0
        DISPLAY "ZERO"
    ELSE
        DISPLAY "NONZERO"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ZERO" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZERO"
        DISPLAY "FAIL: want [ZERO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

