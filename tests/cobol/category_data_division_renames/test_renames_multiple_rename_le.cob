*> vybe-test: cobol/category_data_division_renames/test_renames_multiple_rename_levels
*> origin: languages/cobol/tests/cobol/test_category_data_division_renames.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
01 WS-VYBE-I PIC 9(4) VALUE 0. 01 PARENT. 05 G1 PIC X VALUE 'A'. 05 G2 PIC X VALUE 'B'. 05 G3 PIC X VALUE 'C'. 66 R1 RENAMES G1 THRU G2. 66 R2 RENAMES G2. PROCEDURE DIVISION. DISPLAY R1 R2.
    ADD 1 TO WS-VYBE-I
    MOVE SPACES TO WS-VYBE-L
    STRING R1 DELIMITED SIZE R2 DELIMITED SIZE INTO WS-VYBE-L
    EVALUATE WS-VYBE-I
        WHEN 1
            IF WS-VYBE-L NOT = "AB"
                DISPLAY "FAIL at 1 want [AB] got [" WS-VYBE-L "]"
                MOVE 1 TO RETURN-CODE
                RAISE EXCEPTION EC-PROGRAM
            END-IF
        WHEN 2
            IF WS-VYBE-L NOT = "B"
                DISPLAY "FAIL at 2 want [B] got [" WS-VYBE-L "]"
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

