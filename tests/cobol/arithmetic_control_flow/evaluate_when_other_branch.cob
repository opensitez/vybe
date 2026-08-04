*> vybe-test: cobol/arithmetic_control_flow/evaluate_when_other_branch
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE X
        WHEN 1 DISPLAY "A"
        WHEN 2 DISPLAY "B"
        WHEN OTHER DISPLAY "Z"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Z"
        DISPLAY "FAIL: want [Z] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

