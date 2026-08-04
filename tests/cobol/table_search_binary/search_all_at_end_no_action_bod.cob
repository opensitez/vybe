*> vybe-test: cobol/table_search_binary/search_all_at_end_no_action_body_compiles
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 V PIC 9(4) OCCURS 10 TIMES ASCENDING KEY V INDEXED BY VI.
PROCEDURE DIVISION.
    SEARCH ALL V
        AT END CONTINUE
        WHEN V(VI) = 7
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

