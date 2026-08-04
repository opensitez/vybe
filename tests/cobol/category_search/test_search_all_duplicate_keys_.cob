*> vybe-test: cobol/category_search/test_search_all_duplicate_keys_prefers_first_match
*> origin: languages/cobol/tests/cobol/test_category_search.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-ALL-DUP.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 TBL-DATA.
          05 FILLER PIC X(6) VALUE "01A".
          05 FILLER PIC X(6) VALUE "01B".
          05 FILLER PIC X(6) VALUE "02C".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES
             ASCENDING KEY IS KEY-ID
             INDEXED BY IDX.
             10 KEY-ID PIC 9(2).
             10 VAL PIC X(4).
       PROCEDURE DIVISION.
           SEARCH ALL TBL-ENTRY
              WHEN KEY-ID(IDX) = 01
                 DISPLAY VAL(IDX)
           END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING VAL(IDX) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "A"
        DISPLAY "FAIL: want [A] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

