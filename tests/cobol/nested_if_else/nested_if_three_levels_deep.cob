*> vybe-test: cobol/nested_if_else/nested_if_three_levels_deep
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 1.
01 B PIC 9 VALUE 2.
01 C PIC 9 VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A > 0
        IF B > 0
            IF C > 0
                DISPLAY "DEEP TRUE"
            ELSE
                DISPLAY "C FAIL"
            END-IF
        ELSE
            DISPLAY "B FAIL"
        END-IF
    ELSE
        DISPLAY "A FAIL"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "DEEP TRUE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "DEEP TRUE"
        DISPLAY "FAIL: want [DEEP TRUE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

