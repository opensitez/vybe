*> vybe-test: cobol/condition_compound/condition_not_greater
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 3.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NOT N > 5
        DISPLAY "LE"
    ELSE
        DISPLAY "GT"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "LE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LE"
        DISPLAY "FAIL: want [LE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

