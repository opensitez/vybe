*> vybe-test: cobol/table_search_binary/search_linear_found_at_first
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    MOVE "AAA" TO E(1). MOVE "BBB" TO E(2). MOVE "CCC" TO E(3).
    MOVE "DDD" TO E(4). MOVE "EEE" TO E(5).
    SET IX TO 1.
    SEARCH E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = "AAA" DISPLAY "FOUND FIRST"
    END-SEARCH.
    STOP RUN.

