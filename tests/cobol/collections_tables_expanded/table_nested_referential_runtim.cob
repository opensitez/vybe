*> vybe-test: cobol/collections_tables_expanded/table_nested_referential_runtime
*> origin: languages/cobol/tests/cobol/test_collections_tables_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COLS.
   05 WS-ROW OCCURS 2 TIMES.
      10 WS-COL PIC X(2) OCCURS 2 TIMES.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "A1" TO WS-COL(1,1).
    MOVE "A2" TO WS-COL(1,2).
    MOVE "B1" TO WS-COL(2,1).
    MOVE "B2" TO WS-COL(2,2).
    DISPLAY WS-COL(2,1).
    MOVE SPACES TO WS-VYBE-L
    STRING WS-COL(2,1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "B1"
        DISPLAY "FAIL: want [B1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    DISPLAY WS-COL(1,2).
    MOVE SPACES TO WS-VYBE-L
    STRING WS-COL(1,2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A2"
        DISPLAY "FAIL: want [A2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

