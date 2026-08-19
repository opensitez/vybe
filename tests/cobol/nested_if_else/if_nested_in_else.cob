*> vybe-test: cobol/nested_if_else/if_nested_in_else
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N > 5
        DISPLAY "HIGH"
    ELSE
        IF N > 0
            DISPLAY "LOW"
        ELSE
            DISPLAY "ZERO"
        END-IF
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "HIGH" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HIGH"
        DISPLAY "FAIL: want [HIGH] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

