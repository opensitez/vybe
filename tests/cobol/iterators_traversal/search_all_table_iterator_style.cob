*> vybe-test: cobol/iterators_traversal/search_all_table_iterator_style_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_traversal.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 5 TIMES ASCENDING KEY IS WS-KEY INDEXED BY WS-IDX.
      10 WS-KEY PIC 9(3).
PROCEDURE DIVISION.
    SEARCH ALL WS-ENTRY
        WHEN WS-KEY(WS-IDX) = 200 DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

