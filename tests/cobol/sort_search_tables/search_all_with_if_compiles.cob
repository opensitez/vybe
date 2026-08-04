*> vybe-test: cobol/sort_search_tables/search_all_with_if_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 5 TIMES ASCENDING KEY K INDEXED BY I.
      10 K PIC 9(3).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 1 TO K(1).
    MOVE 2 TO K(2).
    IF 1 = 1
        SEARCH ALL E WHEN K(I) = 1 DISPLAY "YES"
    END-SEARCH
    END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "YES" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "YES"
        DISPLAY "FAIL: want [YES] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

