*> vybe-test: cobol/sort_search_tables/search_loop_wrapper_compiles
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 0.
01 T.
   05 E OCCURS 5 TIMES INDEXED BY I.
      10 K PIC X(3).
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE "A" TO K(1).
    MOVE "A" TO K(2).
    SET I TO 1.
    PERFORM UNTIL N >= 2
        ADD 1 TO N
        SEARCH E WHEN K(I) = "A" DISPLAY N END-SEARCH
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING N DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "2"
                DISPLAY "FAIL at 1 want [2] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "2"
                DISPLAY "FAIL at 2 want [2] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 2 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    IF WS-VYBE-I NOT = 2
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 2"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

