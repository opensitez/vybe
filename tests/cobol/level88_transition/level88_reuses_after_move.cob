*> vybe-test: cobol/level88_transition/level88_reuses_after_move
*> origin: languages/cobol/tests/cobol/test_level88_transition.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X VALUE "N".
    88 FLAGGED VALUE "Y".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "Y" TO S.
    IF FLAGGED
        DISPLAY "SET"
    ELSE
        DISPLAY "UNSET"
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "SET" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "SET"
        DISPLAY "FAIL: want [SET] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

