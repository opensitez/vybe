*> vybe-test: cobol/category_search/test_search_linear_not_found
*> origin: languages/cobol/tests/cobol/test_category_search.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-NOT-FOUND.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(3) VALUE "AAA".
          05 FILLER PIC X(3) VALUE "BBB".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 2 TIMES INDEXED BY IDX.
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN VAL(IDX) = "CCC"
                 DISPLAY "FOUND"
           END-SEARCH.
           STOP RUN.

