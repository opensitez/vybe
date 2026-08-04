*> vybe-test: cobol/condition_compound/condition_equal_spaces_alphanumeric
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE SPACES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF S = SPACES
        DISPLAY "BLANK"
    ELSE
        DISPLAY "NOT BLANK"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "BLANK" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "BLANK"
        DISPLAY "FAIL: want [BLANK] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

