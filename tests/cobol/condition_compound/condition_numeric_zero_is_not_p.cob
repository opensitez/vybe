*> vybe-test: cobol/condition_compound/condition_numeric_zero_is_not_positive
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF N > 0
        DISPLAY "POS"
    ELSE
        DISPLAY "ZERO OR NEG"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "POS" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "POS"
        DISPLAY "FAIL: want [POS] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

