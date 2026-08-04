*> vybe-test: cobol/arrays_tables_indexing/table_search_found_runtime
*> origin: languages/cobol/tests/cobol/test_arrays_tables_indexing.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 3 TIMES ASCENDING KEY IS K INDEXED BY I.
      10 K PIC 9(2).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE 10 TO K(1).
    MOVE 20 TO K(2).
    MOVE 30 TO K(3).
    SEARCH ALL E
        WHEN K(I) = 20
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

