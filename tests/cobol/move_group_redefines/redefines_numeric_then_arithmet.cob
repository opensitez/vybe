*> vybe-test: cobol/move_group_redefines/redefines_numeric_then_arithmetic
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 BUF PIC X(4) VALUE "0099".
01 N REDEFINES BUF PIC 9(4).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    ADD 1 TO N.
    DISPLAY N.
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0100"
        DISPLAY "FAIL: want [0100] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

