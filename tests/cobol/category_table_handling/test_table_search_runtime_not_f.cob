*> vybe-test: cobol/category_table_handling/test_table_search_runtime_not_found
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-SRCH-NF.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 WS.
          05 ELEM OCCURS 3 TIMES INDEXED BY I PIC 9(2).
       PROCEDURE DIVISION.
           MOVE 11 TO ELEM(1).
           MOVE 22 TO ELEM(2).
           MOVE 33 TO ELEM(3).
           SEARCH WS
               AT END DISPLAY "NONE"
               WHEN ELEM(I) = 99 DISPLAY "FOUND"
           END-SEARCH.
           STOP RUN.

