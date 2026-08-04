*> vybe-test: cobol/table_search_binary/search_all_with_action_on_found
*> origin: languages/cobol/tests/cobol/test_table_search_binary.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 PRODUCTS.
   05 PROD OCCURS 100 TIMES ASCENDING KEY PROD-CODE INDEXED BY PI.
      10 PROD-CODE PIC 9(5).
      10 PROD-DESC PIC X(20).
      10 PROD-PRICE PIC 9(7)V99.
PROCEDURE DIVISION.
    SEARCH ALL PROD
        AT END
            DISPLAY "NOT FOUND"
        WHEN PROD-CODE(PI) = 10001
            DISPLAY PROD-DESC(PI)
    END-SEARCH.
    STOP RUN.

