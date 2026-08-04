*> vybe-test: cobol/category_data_editing_advanced/test_edit_z_suppress_partial_zeros
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC ZZ9. 01 Y PIC 999 VALUE 0. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE X DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[  0]"
        DISPLAY "FAIL: want [[  0]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

