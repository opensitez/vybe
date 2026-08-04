*> vybe-test: cobol/occurs_indexed_by/occurs_search_linear_found
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    MOVE "AAA" TO E(1).
    MOVE "BBB" TO E(2).
    MOVE "CCC" TO E(3).
    SET IX TO 1.
    SEARCH E
        AT END
            DISPLAY "NOT FOUND"
        WHEN E(IX) = "BBB"
            DISPLAY "FOUND BBB"
    END-SEARCH.
    STOP RUN.

