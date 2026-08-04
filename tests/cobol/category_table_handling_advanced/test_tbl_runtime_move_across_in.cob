*> vybe-test: cobol/category_table_handling_advanced/test_tbl_runtime_move_across_indexes
*> origin: languages/cobol/tests/cobol/test_category_table_handling_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 S. 05 A OCCURS 2 TIMES PIC 9. 01 D. 05 B OCCURS 2 TIMES PIC 9. PROCEDURE DIVISION. MOVE 4 TO A(1) MOVE 5 TO A(2) MOVE A(1) TO B(2) MOVE A(2) TO B(1) DISPLAY B(1) DISPLAY B(2).
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING B(1) DELIMITED SIZE DISPLAY DELIMITED SIZE B(2) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "5"
                DISPLAY "FAIL at 1 want [5] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "4"
                DISPLAY "FAIL at 2 want [4] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 2 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. STOP RUN.
    IF WS-VYBE-I NOT = 2
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 2"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

