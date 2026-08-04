*> vybe-test: cobol/condition_compound/condition_not_less_means_ge
*> origin: languages/cobol/tests/cobol/test_condition_compound.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 5.
01 B PIC 9 VALUE 5.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF NOT A < B
        DISPLAY "GE"
    ELSE
        DISPLAY "LT"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "GE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "GE"
        DISPLAY "FAIL: want [GE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

