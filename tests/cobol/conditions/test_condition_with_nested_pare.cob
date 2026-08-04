*> vybe-test: cobol/conditions/test_condition_with_nested_parentheses_precedence
*> origin: languages/cobol/tests/cobol/test_conditions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9 VALUE 1.
01 WS-B PIC 9 VALUE 8.
01 WS-C PIC 9 VALUE 9.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    IF (WS-A > 0 AND WS-B > 5) OR WS-C > 10
        DISPLAY "COND"
    ELSE
        DISPLAY "NONE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "COND" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "COND"
        DISPLAY "FAIL: want [COND] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

