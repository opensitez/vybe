*> vybe-test: cobol/iterators_generators_control/iter_search_all_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_generators_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.
      10 K PIC X(3).
PROCEDURE DIVISION.
    SEARCH ALL E
        WHEN K(I) = "ABC" DISPLAY "F"
    END-SEARCH.
    STOP RUN.

