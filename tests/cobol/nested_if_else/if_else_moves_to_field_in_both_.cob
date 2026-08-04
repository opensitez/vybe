*> vybe-test: cobol/nested_if_else/if_else_moves_to_field_in_both_branches
*> origin: languages/cobol/tests/cobol/test_nested_if_else.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 COND PIC 9 VALUE 0.
01 RESULT PIC X(4) VALUE "----".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF COND = 1
        MOVE "TRUE" TO RESULT
    ELSE
        MOVE "FALS" TO RESULT
    END-IF.
    DISPLAY RESULT.
    MOVE SPACES TO WS-VYBE-L
    STRING RESULT DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FALS"
        DISPLAY "FAIL: want [FALS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

