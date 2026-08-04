*> vybe-test: cobol/table_subscript_index/table_binary_search_at_end
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 SORTED-T.
   05 ENTRY OCCURS 5 TIMES ASCENDING KEY ENTRY INDEXED BY S-IX.
      10 ENTRY PIC 9(2).
PROCEDURE DIVISION.
    SEARCH ALL ENTRY
        AT END
            DISPLAY "NOT IN TABLE"
        WHEN ENTRY(S-IX) = 99
            DISPLAY "FOUND 99"
    END-SEARCH.
    STOP RUN.

