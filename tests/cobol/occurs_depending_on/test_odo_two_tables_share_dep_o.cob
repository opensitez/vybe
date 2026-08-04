*> vybe-test: cobol/occurs_depending_on/test_odo_two_tables_share_dep_on_count
*> origin: languages/cobol/tests/cobol/test_occurs_depending_on.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 99 VALUE 2.
01 WS-TABLE-STR.
   05 WS-LABEL PIC X(2) OCCURS 1 TO 5 TIMES DEPENDING ON WS-COUNT.
01 WS-TABLE-NUM.
   05 WS-NUMBER PIC 9(2) OCCURS 1 TO 5 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.

    MOVE "AA" TO WS-LABEL(1).
    MOVE "BB" TO WS-LABEL(2).
    MOVE "CC" TO WS-LABEL(3).
    MOVE 10 TO WS-NUMBER(1).
    MOVE 20 TO WS-NUMBER(2).
    MOVE 30 TO WS-NUMBER(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-LABEL(WS-I)
        DISPLAY WS-NUMBER(WS-I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-LABEL(WS-I) DELIMITED SIZE DISPLAY DELIMITED SIZE WS-NUMBER(WS-I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "AA"
                DISPLAY "FAIL at 1 want [AA] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 2 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "BB"
                DISPLAY "FAIL at 3 want [BB] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 4 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "AA"
                DISPLAY "FAIL at 5 want [AA] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 6 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "BB"
                DISPLAY "FAIL at 7 want [BB] got [" WS-VYBE-L "]"
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
            IF WS-VYBE-L NOT = "CC"
                DISPLAY "FAIL at 9 want [CC] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 10
            IF WS-VYBE-L NOT = "30"
                DISPLAY "FAIL at 10 want [30] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 10 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    MOVE 3 TO WS-COUNT.
    MOVE "CC" TO WS-LABEL(3).
    MOVE 30 TO WS-NUMBER(3).
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > WS-COUNT
        DISPLAY WS-LABEL(WS-I)
        DISPLAY WS-NUMBER(WS-I)
    END-PERFORM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-LABEL(WS-I) DELIMITED SIZE DISPLAY DELIMITED SIZE WS-NUMBER(WS-I) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "AA"
                DISPLAY "FAIL at 1 want [AA] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 2 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "BB"
                DISPLAY "FAIL at 3 want [BB] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 4
            IF WS-VYBE-L NOT = "20"
                DISPLAY "FAIL at 4 want [20] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 5
            IF WS-VYBE-L NOT = "AA"
                DISPLAY "FAIL at 5 want [AA] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 6
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 6 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 7
            IF WS-VYBE-L NOT = "BB"
                DISPLAY "FAIL at 7 want [BB] got [" WS-VYBE-L "]"
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
            IF WS-VYBE-L NOT = "CC"
                DISPLAY "FAIL at 9 want [CC] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 10
            IF WS-VYBE-L NOT = "30"
                DISPLAY "FAIL at 10 want [30] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 10 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.

    IF WS-VYBE-I NOT = 10
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 10"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

