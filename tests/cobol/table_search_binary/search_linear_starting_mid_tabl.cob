*> vybe-test: cobol/table_search_binary/search_linear_starting_mid_table
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
    SET IX TO 3.
    SEARCH E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = "CCC" DISPLAY "FOUND MID"
    END-SEARCH.
    STOP RUN.

