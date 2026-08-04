*> vybe-test: cobol/conditions_extended/nested_if_else_runtime
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 8.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF WS-A > 0
        IF WS-A < 10
            DISPLAY "MID"
        ELSE
            DISPLAY "BIG"
        END-IF
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "MID" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "MID"
        DISPLAY "FAIL: want [MID] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

