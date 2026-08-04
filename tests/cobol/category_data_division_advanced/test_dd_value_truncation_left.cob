*> vybe-test: cobol/category_data_division_advanced/test_dd_value_truncation_left
*> origin: languages/cobol/tests/cobol/test_category_data_division_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC 9(2) VALUE 123. PROCEDURE DIVISION. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "23"
        DISPLAY "FAIL: want [23] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

