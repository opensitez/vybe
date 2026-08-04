*> vybe-test: cobol/display_formatting/display_boolean_condition_result
*> origin: languages/cobol/tests/cobol/test_display_formatting.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF A > B
        DISPLAY "GREATER"
    ELSE
        DISPLAY "NOT GREATER"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "GREATER" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "GREATER"
        DISPLAY "FAIL: want [GREATER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

