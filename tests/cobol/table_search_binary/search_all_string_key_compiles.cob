*> vybe-test: cobol/table_search_binary/search_all_string_key_compiles
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DICT.
   05 ENTRY OCCURS 20 TIMES ASCENDING KEY WORD INDEXED BY DI.
      10 WORD PIC X(10).
      10 DEFN PIC X(40).
PROCEDURE DIVISION.
    SEARCH ALL ENTRY
        AT END DISPLAY "MISSING"
        WHEN WORD(DI) = "COBOL     "
            DISPLAY DEFN(DI)
    END-SEARCH.
    STOP RUN.

