*> vybe-test: cobol/move_group_redefines/redefines_used_in_evaluate
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 BASE PIC X(2) VALUE "01".
01 CODE-NUM REDEFINES BASE PIC 9(2).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    EVALUATE CODE-NUM
        WHEN 1
            DISPLAY "ONE"
        WHEN OTHER
            DISPLAY "OTHER"
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "ONE" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ONE"
        DISPLAY "FAIL: want [ONE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

