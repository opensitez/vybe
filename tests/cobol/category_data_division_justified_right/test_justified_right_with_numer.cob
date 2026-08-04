*> vybe-test: cobol/category_data_division_justified_right/test_justified_right_with_numeric_field
*> origin: languages/cobol/tests/cobol/test_category_data_division_justified_right.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 N PIC 999 JUSTIFIED RIGHT. PROCEDURE DIVISION. MOVE 5 TO N. DISPLAY '[' N ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE N DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[  5]"
        DISPLAY "FAIL: want [[  5]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

