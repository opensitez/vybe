*> vybe-test: cobol/scope_terminators_nesting/scope_end_evaluate_runs_when_other
*> origin: languages/cobol/tests/cobol/test_scope_terminators_nesting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 7.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE N
        WHEN 1
            DISPLAY "ONE"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OTHER"
        DISPLAY "FAIL: want [OTHER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

