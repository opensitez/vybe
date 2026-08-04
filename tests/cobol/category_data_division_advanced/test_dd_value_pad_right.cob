*> vybe-test: cobol/category_data_division_advanced/test_dd_value_pad_right
*> origin: languages/cobol/tests/cobol/test_category_data_division_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC X(4) VALUE 'A'. PROCEDURE DIVISION. DISPLAY '[' R ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE R DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[A   ]"
        DISPLAY "FAIL: want [[A   ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

