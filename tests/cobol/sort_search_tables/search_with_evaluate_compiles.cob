*> vybe-test: cobol/sort_search_tables/search_with_evaluate_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 5 TIMES INDEXED BY I.
      10 K PIC X(3).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE "A" TO K(1).
    MOVE "B" TO K(2).
    EVALUATE TRUE
    WHEN 1 = 1 SEARCH E WHEN K(I) = "B" DISPLAY "EVAL"
    END-SEARCH
    END-EVALUATE.
    MOVE SPACES TO WS-VYBE-L
    STRING "EVAL" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "EVAL"
        DISPLAY "FAIL: want [EVAL] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

