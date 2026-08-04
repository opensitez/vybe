*> vybe-test: cobol/arrays_tables_indexing/table_search_compiles
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 5 TIMES INDEXED BY I.
      10 K PIC X(3).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "AAA" TO K(1).
    MOVE "BBB" TO K(2).
    MOVE "CCC" TO K(3).
    SEARCH E
        WHEN K(I) = "BBB" DISPLAY "FOUND"
    END-SEARCH.
    MOVE SPACES TO WS-VYBE-L
    STRING "FOUND" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "FOUND"
        DISPLAY "FAIL: want [FOUND] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

