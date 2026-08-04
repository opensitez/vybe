*> vybe-test: cobol/collections_dictionaries/map_search_all_by_key_compiles
*> origin: languages/cobol/tests/cobol/test_collections_dictionaries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.
      10 K PIC X(5).
PROCEDURE DIVISION.
    SEARCH ALL E
        WHEN K(I) = "A" DISPLAY "F"
    END-SEARCH.
    STOP RUN.

