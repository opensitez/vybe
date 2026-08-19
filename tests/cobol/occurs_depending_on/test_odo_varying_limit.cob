*> vybe-test: cobol/occurs_depending_on/test_odo_varying_limit
*> origin: languages/cobol/tests/cobol/test_occurs_depending_on.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 99 VALUE 3.
01 WS-TABLE.
   05 WS-ITEM PIC 9(3) OCCURS 1 TO 10 TIMES DEPENDING ON WS-COUNT.
01 WS-I PIC 99 VALUE 1.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.

    MOVE 100 TO WS-ITEM(1).
    MOVE 200 TO WS-ITEM(2).
    MOVE 300 TO WS-ITEM(3).
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
            IF WS-VYBE-L NOT = "200"
                DISPLAY "FAIL at 2 want [200] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "300"
                DISPLAY "FAIL at 3 want [300] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 3 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.

    IF WS-VYBE-I NOT = 3
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 3"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

