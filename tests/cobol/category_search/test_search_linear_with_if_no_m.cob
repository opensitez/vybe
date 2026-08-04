*> vybe-test: cobol/category_search/test_search_linear_with_if_no_match
*> origin: languages/cobol/tests/cobol/test_category_search.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-IF-NOMATCH.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X VALUE "1".
          05 FILLER PIC X VALUE "2".
          05 FILLER PIC X VALUE "3".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES INDEXED BY IDX.
             10 VAL PIC X.
       01 RESULT PIC X VALUE "N".
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY
              AT END IF RESULT = "N" DISPLAY "NO" END-IF
              WHEN VAL(IDX) = "9"
                 MOVE "Y" TO RESULT
                 DISPLAY RESULT
           END-SEARCH.
           STOP RUN.

