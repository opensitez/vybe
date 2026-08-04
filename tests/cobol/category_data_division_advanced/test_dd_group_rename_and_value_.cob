*> vybe-test: cobol/category_data_division_advanced/test_dd_group_rename_and_value_move
*> origin: languages/cobol/tests/cobol/test_category_data_division_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 G. 05 A PIC X VALUE 'A'. 05 B PIC X VALUE 'B'. 05 C PIC X VALUE 'C'. 66 R RENAMES A THRU B. PROCEDURE DIVISION. MOVE 'ZZ' TO R DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZZ"
        DISPLAY "FAIL: want [ZZ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

