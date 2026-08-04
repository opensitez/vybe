*> vybe-test: cobol/occurs_indexed_by/occurs_search_all_requires_ascending_key_compiles
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(3) OCCURS 10 TIMES ASCENDING KEY E INDEXED BY IX.
PROCEDURE DIVISION.
    SEARCH ALL E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = 5
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

