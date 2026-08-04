*> vybe-test: cobol/category_data_division_synchronized/test_sync_grouping_with_right_justified_child
*> origin: languages/cobol/tests/cobol/test_category_data_division_synchronized.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 G. 05 A PIC X(4) VALUE 'ZZ' SYNCHRONIZED RIGHT. 05 B PIC X(4) VALUE 'X' SYNCHRONIZED LEFT. PROCEDURE DIVISION. DISPLAY '[' A ']' DISPLAY '[' B ']'.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE A DELIMITED SIZE ']' DELIMITED SIZE DISPLAY DELIMITED SIZE '[' DELIMITED SIZE B DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "[ZZ  ]"
                DISPLAY "FAIL at 1 want [[ZZ  ]] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "[X   ]"
                DISPLAY "FAIL at 2 want [[X   ]] got [" WS-VYBE-L "]"
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

