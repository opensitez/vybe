*> vybe-test: cobol/conditions/test_condition_nested_if_else_chain
*> origin: languages/cobol/tests/cobol/test_conditions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 2.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF WS-A > 5
        DISPLAY "BIG"
    ELSE
        IF WS-A = 2
            DISPLAY "TWO"
        ELSE
            DISPLAY "OTHER"
        END-IF
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BIG" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BIG"
        DISPLAY "FAIL: want [BIG] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

