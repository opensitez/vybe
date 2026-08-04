*> vybe-test: cobol/move_group_redefines/move_group_from_group_different_sizes_pads
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SRC.
   05 A PIC X(3) VALUE "ABC".
01 DST.
   05 B PIC X(6) VALUE "XXXXXX".
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE SRC TO DST.
    DISPLAY B.
    MOVE SPACES TO WS-VYBE-L
    STRING B DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC   "
        DISPLAY "FAIL: want [ABC   ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

