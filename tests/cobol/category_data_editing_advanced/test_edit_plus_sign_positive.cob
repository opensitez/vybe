*> vybe-test: cobol/category_data_editing_advanced/test_edit_plus_sign_positive
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC +++. 01 Y PIC S999 VALUE 5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE X DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[ +5]"
        DISPLAY "FAIL: want [[ +5]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

