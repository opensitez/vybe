*> vybe-test: cobol/category_search/test_search_multiple_when
*> origin: languages/cobol/tests/cobol/test_category_search.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-MULTI-WHEN.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X VALUE "A".
          05 FILLER PIC X VALUE "B".
          05 FILLER PIC X VALUE "C".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES INDEXED BY IDX.
             10 VAL PIC X.
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN VAL(IDX) = "X" DISPLAY "FOUND X"
              WHEN VAL(IDX) = "B" DISPLAY "FOUND B"
           END-SEARCH.
           STOP RUN.

