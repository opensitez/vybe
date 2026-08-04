*> vybe-test: cobol/table_search_binary/search_linear_multiple_search_calls
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X(2) OCCURS 5 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    MOVE "A1" TO E(1). MOVE "B2" TO E(2). MOVE "C3" TO E(3).
    MOVE "D4" TO E(4). MOVE "E5" TO E(5).
    SET IX TO 1.
    SEARCH E
        AT END DISPLAY "MISS1"
        WHEN E(IX) = "C3" DISPLAY "HIT C3"
    END-SEARCH.
    SET IX TO 1.
    SEARCH E
        AT END DISPLAY "MISS2"
        WHEN E(IX) = "E5" DISPLAY "HIT E5"
    END-SEARCH.
    STOP RUN.

