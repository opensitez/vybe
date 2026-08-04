*> vybe-test: cobol/tables/test_table_linear_search_with_at_end
*> origin: languages/cobol/tests/cobol/test_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 4 TIMES
      INDEXED BY WS-IDX.
      10 WS-KEY PIC X(2).
PROCEDURE DIVISION.

    MOVE "A" TO WS-KEY(1).
    MOVE "B" TO WS-KEY(2).
    MOVE "C" TO WS-KEY(3).
    SEARCH WS-ENTRY
        AT END DISPLAY 'NONE'
        WHEN WS-KEY(WS-IDX) = "B" DISPLAY 'FOUND'
    END-SEARCH.
    STOP RUN.

