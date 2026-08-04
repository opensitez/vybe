*> vybe-test: cobol/move_group_redefines/redefines_write_via_redefine_reads_original
*> origin: languages/cobol/tests/cobol/test_move_group_redefines.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 BUF PIC X(4) VALUE "XXXX".
01 ALIAS REDEFINES BUF PIC X(4).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "COBO" TO ALIAS.
    DISPLAY BUF.
    MOVE SPACES TO WS-VYBE-L
    STRING BUF DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "COBO"
        DISPLAY "FAIL: want [COBO] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

