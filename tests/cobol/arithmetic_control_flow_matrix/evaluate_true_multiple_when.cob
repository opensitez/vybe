*> vybe-test: cobol/arithmetic_control_flow_matrix/evaluate_true_multiple_when
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 2.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE TRUE
        WHEN X = 1 DISPLAY "A"
        WHEN X = 2 DISPLAY "B"
        WHEN OTHER DISPLAY "Z"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "A" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

