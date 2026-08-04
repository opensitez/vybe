*> vybe-test: cobol/table_subscript_index/table_ascending_initialized_order
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(3) OCCURS 5 TIMES.
01 I PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE 10 TO E(1). MOVE 20 TO E(2). MOVE 30 TO E(3).
    MOVE 40 TO E(4). MOVE 50 TO E(5).
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 5
        DISPLAY E(I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING E(I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "010"
                DISPLAY "FAIL at 1 want [010] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "020"
                DISPLAY "FAIL at 2 want [020] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "030"
                DISPLAY "FAIL at 3 want [030] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "040"
                DISPLAY "FAIL at 4 want [040] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "050"
                DISPLAY "FAIL at 5 want [050] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 5 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    IF WS-VYBE-I NOT = 5
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 5"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

