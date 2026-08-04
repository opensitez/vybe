*> vybe-test: cobol/sort_search_tables/sort_then_search_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC X(10).
01 K PIC 9(5).
01 T.
   05 E OCCURS 5 TIMES ASCENDING KEY K2 INDEXED BY I.
      10 K2 PIC 9(5).
PROCEDURE DIVISION.
    SORT F ON ASCENDING KEY K.
    SEARCH ALL E WHEN K2(I) = 10 DISPLAY "F" END-SEARCH.
    STOP RUN.

