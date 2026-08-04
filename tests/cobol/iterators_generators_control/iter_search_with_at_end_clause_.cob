*> vybe-test: cobol/iterators_generators_control/iter_search_with_at_end_clause_compiles
*> origin: languages/cobol/tests/cobol/test_iterators_generators_control.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 3 TIMES INDEXED BY I.
      10 K PIC X(3).
PROCEDURE DIVISION.
    SEARCH E
        AT END DISPLAY "NONE"
        WHEN K(I) = "ABC" DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

