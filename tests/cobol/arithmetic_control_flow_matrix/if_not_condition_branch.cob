*> vybe-test: cobol/arithmetic_control_flow_matrix/if_not_condition_branch
*> origin: languages/cobol/tests/cobol/test_arithmetic_control_flow_matrix.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NOT A = 1
        DISPLAY "N1"
    ELSE
        DISPLAY "N0"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "N1" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "N1"
        DISPLAY "FAIL: want [N1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

