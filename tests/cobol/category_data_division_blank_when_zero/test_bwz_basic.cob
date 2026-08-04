*> vybe-test: cobol/category_data_division_blank_when_zero/test_bwz_basic
*> origin: languages/cobol/tests/cobol/test_category_data_division_blank_when_zero.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC 9(3) VALUE 0 BLANK WHEN ZERO. PROCEDURE DIVISION. DISPLAY '[' V ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE V DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[   ]"
        DISPLAY "FAIL: want [[   ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

