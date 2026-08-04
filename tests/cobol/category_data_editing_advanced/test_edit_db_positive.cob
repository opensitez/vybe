*> vybe-test: cobol/category_data_editing_advanced/test_edit_db_positive
*> origin: languages/cobol/tests/cobol/test_category_data_editing_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC 99DB. 01 Y PIC S99 VALUE 5. PROCEDURE DIVISION. MOVE Y TO X. DISPLAY '[' X ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE X DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[05  ]"
        DISPLAY "FAIL: want [[05  ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

