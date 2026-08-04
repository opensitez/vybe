*> vybe-test: cobol/collections_tables_expanded/table_search_not_found_runtime
*> origin: languages/cobol/tests/cobol/test_collections_tables_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 3 TIMES ASCENDING KEY IS WS-KEY INDEXED BY WS-IDX.
      10 WS-KEY PIC 9(3).
PROCEDURE DIVISION.
    MOVE 10 TO WS-KEY(1).
    MOVE 20 TO WS-KEY(2).
    MOVE 30 TO WS-KEY(3).
    SEARCH WS-ENTRY
        AT END DISPLAY "NONE"
        WHEN WS-KEY(WS-IDX) = 99 DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

