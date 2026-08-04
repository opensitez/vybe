*> vybe-test: cobol/category_search/test_search_linear_varying
*> origin: languages/cobol/tests/cobol/test_category_search.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-VARYING.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X VALUE "X".
          05 FILLER PIC X VALUE "Y".
          05 FILLER PIC X VALUE "Z".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES INDEXED BY IDX.
             10 VAL PIC X.
       01 TRACKER PIC 9 VALUE 1.
       PROCEDURE DIVISION.
           SET IDX TO 1.
           SEARCH TBL-ENTRY VARYING TRACKER
              AT END DISPLAY "NOT FOUND"
              WHEN VAL(IDX) = "Z"
                 DISPLAY TRACKER
           END-SEARCH.
           STOP RUN.

