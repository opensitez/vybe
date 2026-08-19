*> vybe-test: cobol/occurs_depending_on/test_odo_count_resize_cycles
*> origin: languages/cobol/tests/cobol/test_occurs_depending_on.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE.
   05 WS-ENTRY PIC 9(2) OCCURS 1 TO 6 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.

    MOVE 11 TO WS-ENTRY(1).
    MOVE 22 TO WS-ENTRY(2).
    MOVE 33 TO WS-ENTRY(3).
    MOVE 44 TO WS-ENTRY(4).
    MOVE 55 TO WS-ENTRY(5).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ENTRY(WS-I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "33"
                DISPLAY "FAIL at 1 want [33] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "22"
                DISPLAY "FAIL at 2 want [22] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "55"
                DISPLAY "FAIL at 3 want [55] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "11"
                DISPLAY "FAIL at 4 want [11] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "22"
                DISPLAY "FAIL at 5 want [22] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "33"
                DISPLAY "FAIL at 6 want [33] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "44"
                DISPLAY "FAIL at 7 want [44] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 7 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    MOVE 1 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ENTRY(WS-I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "11"
                DISPLAY "FAIL at 1 want [11] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "22"
                DISPLAY "FAIL at 2 want [22] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "55"
                DISPLAY "FAIL at 3 want [55] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "11"
                DISPLAY "FAIL at 4 want [11] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "22"
                DISPLAY "FAIL at 5 want [22] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "33"
                DISPLAY "FAIL at 6 want [33] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "44"
                DISPLAY "FAIL at 7 want [44] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 7 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    MOVE 4 TO WS-COUNT.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-ENTRY(WS-I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ENTRY(WS-I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "11"
                DISPLAY "FAIL at 1 want [11] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "22"
                DISPLAY "FAIL at 2 want [22] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "55"
                DISPLAY "FAIL at 3 want [55] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "11"
                DISPLAY "FAIL at 4 want [11] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "22"
                DISPLAY "FAIL at 5 want [22] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "33"
                DISPLAY "FAIL at 6 want [33] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "44"
                DISPLAY "FAIL at 7 want [44] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 7 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.

    IF WS-VYBE-I NOT = 7
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 7"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

