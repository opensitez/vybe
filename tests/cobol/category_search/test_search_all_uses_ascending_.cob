*> vybe-test: cobol/category_search/test_search_all_uses_ascending_key_order
*> origin: languages/cobol/tests/cobol/test_category_search.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. SEARCH-ALL-ORDER.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 TBL-DATA.
          05 FILLER PIC X(5) VALUE "05PEND".
          05 FILLER PIC X(5) VALUE "02CODE".
          05 FILLER PIC X(5) VALUE "07END ".
       01 TBL REDEFINES TBL-DATA.
          05 TBL-ENTRY OCCURS 3 TIMES
             ASCENDING KEY IS KEY-ID
             INDEXED BY IDX.
             10 KEY-ID PIC 9(2).
             10 VAL PIC X(3).
       PROCEDURE DIVISION.
           SEARCH ALL TBL-ENTRY
              WHEN KEY-ID(IDX) = 07
                 DISPLAY VAL(IDX)
           END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING VAL(IDX) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "END"
        DISPLAY "FAIL: want [END] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

