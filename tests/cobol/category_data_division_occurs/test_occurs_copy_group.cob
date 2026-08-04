*> vybe-test: cobol/category_data_division_occurs/test_occurs_copy_group
*> origin: languages/cobol/tests/cobol/test_category_data_division_occurs.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 SRC. 05 A PIC X(2) OCCURS 2 TIMES VALUE SPACES. 01 DST. 05 B PIC X(2) OCCURS 2 TIMES. PROCEDURE DIVISION. MOVE 'AA' TO A(1) MOVE 'BB' TO A(2) MOVE SRC TO DST DISPLAY B(1) DISPLAY B(2) STOP RUN.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING B(1) DELIMITED SIZE DISPLAY DELIMITED SIZE B(2) DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "AA"
                DISPLAY "FAIL at 1 want [AA] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "BB"
                DISPLAY "FAIL at 2 want [BB] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 2 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE.
    IF WS-VYBE-I NOT = 2
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 2"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

