*> vybe-test: cobol/condition_compound/condition_not_equal
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 7.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NOT N = 5
        DISPLAY "DIFF"
    ELSE
        DISPLAY "SAME"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "DIFF" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "DIFF"
        DISPLAY "FAIL: want [DIFF] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

