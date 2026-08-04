*> vybe-test: cobol/category_data_division_renames/test_renames_parse_two_aliases_same_group
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 R. 05 A PIC X VALUE 'A'. 05 B PIC X VALUE 'B'. 05 C PIC X VALUE 'C'. 66 LEFT RENAMES A THRU B. 66 RIGHT RENAMES B THRU C. PROCEDURE DIVISION. MOVE 'XY' TO LEFT DISPLAY LEFT DISPLAY RIGHT.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING LEFT DELIMITED SIZE DISPLAY DELIMITED SIZE RIGHT DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "XY"
                DISPLAY "FAIL at 1 want [XY] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "YC"
                DISPLAY "FAIL at 2 want [YC] got [" WS-VYBE-L "]"
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

