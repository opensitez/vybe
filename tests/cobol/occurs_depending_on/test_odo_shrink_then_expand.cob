*> vybe-test: cobol/occurs_depending_on/test_odo_shrink_then_expand
*> origin: languages/cobol/tests/cobol/test_occurs_depending_on.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 99 VALUE 4.
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 1 TO 10 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.

    MOVE 10 TO WS-ITEM(1).
    MOVE 20 TO WS-ITEM(2).
    MOVE 30 TO WS-ITEM(3).
    MOVE 40 TO WS-ITEM(4).
    MOVE 4 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ITEM(WS-I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "000"
                DISPLAY "FAIL at 1 want [000] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "030"
                DISPLAY "FAIL at 2 want [030] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "040"
                DISPLAY "FAIL at 3 want [040] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "40"
                DISPLAY "FAIL at 4 want [40] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 5 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 6 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 7 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 8
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 8 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 9
            IF WS-VYBE-L NOT = "55"
                DISPLAY "FAIL at 9 want [55] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 9 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    MOVE 2 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ITEM(WS-I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 1 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "030"
                DISPLAY "FAIL at 2 want [030] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "040"
                DISPLAY "FAIL at 3 want [040] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "40"
                DISPLAY "FAIL at 4 want [40] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 5 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 6 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 7 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 8
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 8 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 9
            IF WS-VYBE-L NOT = "55"
                DISPLAY "FAIL at 9 want [55] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 9 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    MOVE 3 TO WS-COUNT.
    MOVE 55 TO WS-ITEM(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ITEM(WS-I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ITEM(WS-I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 1 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 2 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "040"
                DISPLAY "FAIL at 3 want [040] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "40"
                DISPLAY "FAIL at 4 want [40] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 5 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 6 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 7 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 8
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 8 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 9
            IF WS-VYBE-L NOT = "55"
                DISPLAY "FAIL at 9 want [55] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 9 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.

    IF WS-VYBE-I NOT = 9
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 9"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

