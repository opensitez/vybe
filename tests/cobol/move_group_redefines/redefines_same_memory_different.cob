*> vybe-test: cobol/move_group_redefines/redefines_same_memory_different_pic
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 BASE PIC X(4) VALUE "1234".
01 RDEF REDEFINES BASE PIC 9(4).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 5678 TO RDEF.
    DISPLAY BASE.
    MOVE SPACES TO WS-VYBE-L
    STRING BASE DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5678"
        DISPLAY "FAIL: want [5678] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

