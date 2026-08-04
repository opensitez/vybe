*> vybe-test: cobol/arrays_tables_indexing/table_search_not_found_runtime
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 4 TIMES ASCENDING KEY K INDEXED BY I.
      10 K PIC 9(2).
PROCEDURE DIVISION.
    MOVE 11 TO K(1).
    MOVE 22 TO K(2).
    MOVE 33 TO K(3).
    MOVE 44 TO K(4).
    SEARCH E
        AT END DISPLAY "NOT-FOUND"
        WHEN K(I) = 99
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

