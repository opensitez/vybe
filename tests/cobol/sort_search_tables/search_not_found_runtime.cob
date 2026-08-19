*> vybe-test: cobol/sort_search_tables/search_not_found_runtime
*> origin: languages/cobol/tests/cobol/test_sort_search_tables.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E OCCURS 3 TIMES ASCENDING KEY K INDEXED BY I.
      10 K PIC 9(3).
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE 1 TO K(1).
    MOVE 2 TO K(2).
    MOVE 3 TO K(3).
    SEARCH ALL E WHEN K(I) = 9 DISPLAY 'FOUND' END-SEARCH.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'FOUND' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "FOUND"
                DISPLAY "FAIL at 1 want [FOUND] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    DISPLAY 'END'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING 'END' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "END"
                DISPLAY "FAIL at 1 want [END] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 1 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    IF WS-VYBE-I NOT = 1
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 1"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

