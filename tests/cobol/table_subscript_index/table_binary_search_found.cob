*> vybe-test: cobol/table_subscript_index/table_binary_search_found
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SORTED-T.
   05 ENTRY OCCURS 10 TIMES ASCENDING KEY ENTRY INDEXED BY S-IX.
      10 ENTRY PIC 9(4).
PROCEDURE DIVISION.
    SEARCH ALL ENTRY
        AT END DISPLAY "NOT FOUND"
        WHEN ENTRY(S-IX) = 5
            DISPLAY "FOUND"
    END-SEARCH.
    STOP RUN.

