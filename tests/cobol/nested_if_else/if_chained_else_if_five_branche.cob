*> vybe-test: cobol/nested_if_else/if_chained_else_if_five_branches
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 MONTH PIC 9(2) VALUE 9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF MONTH = 1
        DISPLAY "JAN"
    ELSE IF MONTH = 2
        DISPLAY "FEB"
    ELSE IF MONTH = 6
        DISPLAY "JUN"
    ELSE IF MONTH = 9
        DISPLAY "SEP"
    ELSE
        DISPLAY "OTHER"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "JAN" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "JAN"
        DISPLAY "FAIL: want [JAN] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

