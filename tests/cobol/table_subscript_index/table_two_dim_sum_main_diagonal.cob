*> vybe-test: cobol/table_subscript_index/table_two_dim_sum_main_diagonal
*> origin: languages/cobol/tests/cobol/test_table_subscript_index.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 M.
   05 ROW OCCURS 3 TIMES.
      10 COL PIC 9 OCCURS 3 TIMES.
01 S PIC 9(3) VALUE 0.
01 I PIC 9 VALUE 0.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0.
PROCEDURE DIVISION.
    MOVE 1 TO COL(1 1). MOVE 2 TO COL(2 2). MOVE 3 TO COL(3 3).
    PERFORM VARYING I FROM 1 BY 1 UNTIL I > 3
        ADD COL(I I) TO S
    END-PERFORM.
    DISPLAY S.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "6"
                DISPLAY "FAIL at 1 want [6] got [" WS-VYBE-L "]"
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

