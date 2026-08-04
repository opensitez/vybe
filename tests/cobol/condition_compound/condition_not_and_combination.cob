*> vybe-test: cobol/condition_compound/condition_not_and_combination
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NOT (X = 1 OR X = 2)
        DISPLAY "OTHER"
    ELSE
        DISPLAY "ONE OR TWO"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "OTHER" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "OTHER"
        DISPLAY "FAIL: want [OTHER] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

