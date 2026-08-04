*> vybe-test: cobol/table_search_binary/search_linear_with_or_condition
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC X(3) OCCURS 5 TIMES INDEXED BY IX.
PROCEDURE DIVISION.
    MOVE "AAA" TO E(1). MOVE "BBB" TO E(2).
    SET IX TO 1.
    SEARCH E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = "AAA" OR E(IX) = "BBB"
            DISPLAY "FOUND A OR B"
    END-SEARCH.
    STOP RUN.

