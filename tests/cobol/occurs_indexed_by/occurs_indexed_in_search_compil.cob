*> vybe-test: cobol/occurs_indexed_by/occurs_indexed_in_search_compiles
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X(3) OCCURS 10 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    SEARCH E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = "ABC"
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

