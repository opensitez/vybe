*> vybe-test: cobol/iterators_traversal/search_table_iterator_style_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_traversal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 5 TIMES INDEXED BY WS-IDX.
      10 WS-KEY PIC X(3).
PROCEDURE DIVISION.
    SEARCH WS-ENTRY
        AT END DISPLAY "NONE"
        WHEN WS-KEY(WS-IDX) = "ABC" DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

