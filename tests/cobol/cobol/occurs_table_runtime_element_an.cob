*> vybe-test: cobol/cobol/occurs_table_runtime_element_and_iteration
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. TABLES-RUNTIME.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
01 WS-ITEM PIC 9(4) OCCURS 3 TIMES.
01 WS-I PIC 9 VALUE 1.
01 WS-SUM PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE 10 TO WS-ITEM(1).
    MOVE 20 TO WS-ITEM(2).
    MOVE 30 TO WS-ITEM(3).
    DISPLAY WS-ITEM(1).
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-ITEM(1) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "0010"
                DISPLAY "FAIL at 1 want [0010] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "0060"
                DISPLAY "FAIL at 2 want [0060] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 2 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    PERFORM VARYING WS-I FROM 1 BY 1 UNTIL WS-I > 3
        ADD WS-ITEM(WS-I) TO WS-SUM
    END-PERFORM.
    DISPLAY WS-SUM.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING WS-SUM DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "10"
                DISPLAY "FAIL at 1 want [10] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "0060"
                DISPLAY "FAIL at 2 want [0060] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 2 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    STOP RUN.

    IF WS-VYBE-I NOT = 2
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 2"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

