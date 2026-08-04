*> vybe-test: cobol/category_table_handling/test_table_search_runtime_found
*> origin: languages/cobol/tests/cobol/test_category_table_handling.rs

       IDENTIFICATION DIVISION.
       PROGRAM-ID. TBL-SRCH.
       DATA DIVISION.
       WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
       01 WS.
          05 ELEM OCCURS 4 TIMES INDEXED BY I PIC 9(2).
       PROCEDURE DIVISION.
           MOVE 10 TO ELEM(1).
           MOVE 20 TO ELEM(2).
           MOVE 30 TO ELEM(3).
           SEARCH WS
               WHEN ELEM(I) = 20
                   DISPLAY "FOUND"
           END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING "FOUND" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FOUND"
        DISPLAY "FAIL: want [FOUND] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
           STOP RUN.

