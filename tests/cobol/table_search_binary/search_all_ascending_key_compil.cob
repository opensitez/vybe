*> vybe-test: cobol/table_search_binary/search_all_ascending_key_compiles
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 10 TIMES ASCENDING KEY E INDEXED BY IX.
      10 E PIC 9(4).
PROCEDURE DIVISION.
    SEARCH ALL E
        AT END DISPLAY "NOT FOUND"
        WHEN E(IX) = 5
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

