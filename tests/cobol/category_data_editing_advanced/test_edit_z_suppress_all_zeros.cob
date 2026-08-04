*> vybe-test: cobol/category_data_editing_advanced/test_edit_z_suppress_all_zeros
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC ZZZ. 01 Y PIC 999 VALUE 0. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE X DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[   ]"
        DISPLAY "FAIL: want [[   ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

