*> vybe-test: cobol/sort_search_tables/search_with_at_end_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 5 TIMES INDEXED BY I.
      10 K PIC X(3).
PROCEDURE DIVISION.
    MOVE "A" TO K(1).
    SEARCH E AT END DISPLAY "N" WHEN K(I) = "Z" DISPLAY "FOUND" END-SEARCH.
    STOP RUN.

