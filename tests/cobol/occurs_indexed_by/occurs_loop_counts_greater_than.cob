*> vybe-test: cobol/occurs_indexed_by/occurs_loop_counts_greater_than_value
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 E PIC 9(2) OCCURS 10 TIMES INDEXED BY IX.
01 CNT PIC 9(2) VALUE 0.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    PERFORM VARYING IX FROM 1 BY 1 UNTIL IX > 10
        MOVE IX TO E(IX)
    END-PERFORM.
    PERFORM VARYING IX FROM 1 BY 1 UNTIL IX > 10
        IF E(IX) > 5
            ADD 1 TO CNT
        END-IF
    END-PERFORM.
    DISPLAY CNT.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING CNT DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "05"
                DISPLAY "FAIL at 1 want [05] got [" WS-VYBE-L "]"
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

