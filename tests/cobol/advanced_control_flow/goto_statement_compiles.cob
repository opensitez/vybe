*> vybe-test: cobol/advanced_control_flow/goto_statement_compiles
*> origin: languages/cobol/tests/cobol/test_advanced_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    GO TO LABEL-ONE.
LABEL-ONE.
    DISPLAY "DONE".
    MOVE SPACES TO WS-VYBE-L
    STRING "DONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "DONE"
        DISPLAY "FAIL: want [DONE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

