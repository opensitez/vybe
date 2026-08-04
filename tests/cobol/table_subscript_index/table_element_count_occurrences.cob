*> vybe-test: cobol/table_subscript_index/table_element_count_occurrences
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T.
   05 GRADE PIC X OCCURS 10 TIMES.
01 CNT PIC 9(2) VALUE 0.
01 I PIC 9(2) VALUE 0.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE "A" TO GRADE(1). MOVE "B" TO GRADE(2). MOVE "A" TO GRADE(3).
    MOVE "C" TO GRADE(4). MOVE "A" TO GRADE(5).
    MOVE "B" TO GRADE(6). MOVE "D" TO GRADE(7). MOVE "A" TO GRADE(8).
    MOVE "F" TO GRADE(9). MOVE "A" TO GRADE(10).
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 10
        IF GRADE(I) = "A"
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

