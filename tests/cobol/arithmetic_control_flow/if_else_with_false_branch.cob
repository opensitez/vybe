*> vybe-test: cobol/arithmetic_control_flow/if_else_with_false_branch
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF X = 1
        DISPLAY "TRUE"
    ELSE
        DISPLAY "FALSE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "TRUE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FALSE"
        DISPLAY "FAIL: want [FALSE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

