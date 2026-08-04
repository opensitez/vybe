*> vybe-test: cobol/move_group_redefines/redefines_group_can_be_moved_to
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 BASE PIC X(6) VALUE "AABBCC".
01 REDEF REDEFINES BASE.
   05 R1 PIC X(2).
   05 R2 PIC X(2).
   05 R3 PIC X(2).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "XX" TO R1.
    DISPLAY BASE.
    MOVE SPACES TO WS-VYBE-L
    STRING BASE DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "XXBBCC"
        DISPLAY "FAIL: want [XXBBCC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

