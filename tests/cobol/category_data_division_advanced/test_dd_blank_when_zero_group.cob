*> vybe-test: cobol/category_data_division_advanced/test_dd_blank_when_zero_group
*> origin: languages/cobol/tests/cobol/test_category_data_division_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 G BLANK WHEN ZERO. 05 A PIC 9(2) VALUE 0. PROCEDURE DIVISION. DISPLAY '[' G ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE G DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[  ]"
        DISPLAY "FAIL: want [[  ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

