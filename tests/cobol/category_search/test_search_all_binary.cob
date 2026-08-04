*> vybe-test: cobol/category_search/test_search_all_binary
*> origin: languages/cobol/tests/cobol/test_category_search.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-ALL.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
       01 TBL-DATA.
          05 FILLER PIC X(5) VALUE "01AAA".
          05 FILLER PIC X(5) VALUE "02BBB".
          05 FILLER PIC X(5) VALUE "03CCC".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES
             ASCENDING KEY IS KEY-ID
             INDEXED BY IDX.
             10 KEY-ID PIC 9(2).
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SEARCH ALL TBL-ENTRY
              AT END DISPLAY "NOT FOUND"
              WHEN KEY-ID(IDX) = 03
                 DISPLAY VAL(IDX)
           END-SEARCH.
           STOP RUN.

