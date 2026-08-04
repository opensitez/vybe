*> vybe-test: cobol/table_search_binary/search_all_descending_key_compiles
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REVERSE-T.
   05 RE OCCURS 10 TIMES DESCENDING KEY RE INDEXED BY RI.
      10 RE PIC 9(4).
PROCEDURE DIVISION.
    SEARCH ALL RE
        AT END DISPLAY "END"
        WHEN RE(RI) = 1
            DISPLAY "ONE"
    END-SEARCH.
    STOP RUN.

