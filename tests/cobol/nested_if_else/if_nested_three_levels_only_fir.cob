*> vybe-test: cobol/nested_if_else/if_nested_three_levels_only_first_enters
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 0.
01 B PIC 9 VALUE 1.
01 C PIC 9 VALUE 1.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A > 0
        IF B > 0
            IF C > 0
                DISPLAY "ALL"
            END-IF
        END-IF
    ELSE
        DISPLAY "OUTER FALSE"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "ALL" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ALL"
        DISPLAY "FAIL: want [ALL] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

