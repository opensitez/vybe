*> vybe-test: cobol/sort_search_tables/search_with_set_indexing_runtime
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 4 TIMES INDEXED BY I.
      10 K PIC 9(2).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 10 TO K(1).
    MOVE 20 TO K(2).
    MOVE 30 TO K(3).
    SET I TO 2.
    SEARCH E WHEN K(I) = 30 DISPLAY 'LATE' END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING 'LATE' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LATE"
        DISPLAY "FAIL: want [LATE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

