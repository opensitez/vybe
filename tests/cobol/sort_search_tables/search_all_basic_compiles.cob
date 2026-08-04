*> vybe-test: cobol/sort_search_tables/search_all_basic_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.
      10 K PIC X(3).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "AA" TO K(1).
    MOVE "BB" TO K(2).
    MOVE "CC" TO K(3).
    SEARCH ALL E WHEN K(I) = "BB" DISPLAY "FOUND" END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING "FOUND" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FOUND"
        DISPLAY "FAIL: want [FOUND] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

