*> vybe-test: cobol/collections_tables_expanded/table_two_dimensional_access_runtime
*> origin: languages/cobol/tests/cobol/test_collections_tables_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-MATRIX.
   05 WS-ROW OCCURS 2 TIMES.
      10 WS-COL PIC 9 OCCURS 3 TIMES VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 7 TO WS-COL(2,3).
    DISPLAY WS-COL(2,3).
    MOVE SPACES TO WS-VYBE-L
    STRING WS-COL(2,3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "7"
        DISPLAY "FAIL: want [7] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

