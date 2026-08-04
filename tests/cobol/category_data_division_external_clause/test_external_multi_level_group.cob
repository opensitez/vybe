*> vybe-test: cobol/category_data_division_external_clause/test_external_multi_level_group
*> origin: languages/cobol/tests/cobol/test_category_data_division_external_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 E-GRP IS EXTERNAL. 05 A PIC X VALUE 'X'. 05 B PIC X VALUE 'Y'. 05 C PIC 99 VALUE 99. PROCEDURE DIVISION. DISPLAY A DISPLAY B DISPLAY C.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING A DELIMITED SIZE DISPLAY DELIMITED SIZE B DELIMITED SIZE DISPLAY DELIMITED SIZE C DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "X"
                DISPLAY "FAIL at 1 want [X] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "Y"
                DISPLAY "FAIL at 2 want [Y] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 3
            IF WS-VYBE-L NOT = "99"
                DISPLAY "FAIL at 3 want [99] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN OTHER
            DISPLAY "FAIL: more than 3 line(s)"
            MOVE 1 TO RETURN-CODE
            RAISE EXCEPTION EC-PROGRAM
    END-EVALUATE. STOP RUN.
    IF WS-VYBE-I NOT = 3
        DISPLAY "FAIL: " WS-VYBE-I " line(s), wanted 3"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.

