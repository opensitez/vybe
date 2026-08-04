*> vybe-test: cobol/table_subscript_index/table_sequential_search_not_found
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 KEY PIC X(3) OCCURS 5 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    MOVE "AAA" TO KEY(1). MOVE "BBB" TO KEY(2). MOVE "CCC" TO KEY(3).
    MOVE "DDD" TO KEY(4). MOVE "EEE" TO KEY(5).
    SET IX TO 1.
    SEARCH KEY
        AT END DISPLAY "NOT FOUND"
        WHEN KEY(IX) = "ZZZ"
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

