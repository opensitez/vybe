*> vybe-test: cobol/collections_tables_expanded/table_search_statement_runtime
*> origin: languages/cobol/tests/cobol/test_collections_tables_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 10 TIMES INDEXED BY WS-IDX.
      10 WS-KEY PIC X(4).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "ABCD" TO WS-KEY(1).
    MOVE "WXYZ" TO WS-KEY(2).
    MOVE "LMNO" TO WS-KEY(3).
    SEARCH WS-ENTRY
        WHEN WS-KEY(WS-IDX) = "WXYZ" DISPLAY "FOUND"
    END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING "FOUND" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FOUND"
        DISPLAY "FAIL: want [FOUND] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

