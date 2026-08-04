*> vybe-test: cobol/table_search_binary/search_linear_sets_index_on_found
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(2) OCCURS 5 TIMES INDEXED BY IX.
01 FOUND-POS PIC 9 VALUE 0.
PROCEDURE DIVISION.
    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3).
    MOVE 40 TO E(4). MOVE 50 TO E(5).
    SET IX TO 1.
    SEARCH E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = 30
            SET FOUND-POS TO IX
            DISPLAY FOUND-POS
    END-SEARCH.
    STOP RUN.

