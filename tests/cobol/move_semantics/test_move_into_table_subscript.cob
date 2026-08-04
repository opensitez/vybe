*> vybe-test: cobol/move_semantics/test_move_into_table_subscript
*> origin: languages/cobol/tests/cobol/test_move_semantics.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ITEM PIC 9(2) OCCURS 5 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.

    MOVE 99 TO WS-ITEM(3).
    DISPLAY WS-ITEM(3).
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ITEM(3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "99"
        DISPLAY "FAIL: want [99] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

