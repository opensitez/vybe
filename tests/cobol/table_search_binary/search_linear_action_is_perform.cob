*> vybe-test: cobol/table_search_binary/search_linear_action_is_perform
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(3) OCCURS 5 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3).
    SET IX TO 1.
    SEARCH E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = 20
            PERFORM FOUND-ACTION
    END-SEARCH.
    STOP RUN.
FOUND-ACTION.
    DISPLAY "ACTION TAKEN".
    STOP RUN.

