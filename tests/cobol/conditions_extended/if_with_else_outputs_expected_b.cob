*> vybe-test: cobol/conditions_extended/if_with_else_outputs_expected_branch
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF WS-A > 5
        DISPLAY "BIG"
    ELSE
        DISPLAY "SMALL"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BIG" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SMALL"
        DISPLAY "FAIL: want [SMALL] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

