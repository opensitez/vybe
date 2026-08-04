*> vybe-test: cobol/category_search/test_search_linear_basic
*> origin: languages/cobol/tests/cobol/test_category_search.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-LINEAR.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(3) VALUE "AAA".
          05 FILLER PIC X(3) VALUE "BBB".
          05 FILLER PIC X(3) VALUE "CCC".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES INDEXED BY IDX.
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN VAL(IDX) = "BBB"
                 DISPLAY "FOUND " VAL(IDX)
           END-SEARCH.
           STOP RUN.

