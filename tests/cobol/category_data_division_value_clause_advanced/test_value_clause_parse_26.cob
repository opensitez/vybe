*> vybe-test: cobol/category_data_division_value_clause_advanced/test_value_clause_parse_26
*> origin: languages/cobol/tests/cobol/test_category_data_division_value_clause_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(4) VALUE ALL ' '. PROCEDURE DIVISION. IF V = SPACES DISPLAY 'ALLSPACE'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'ALLSPACE' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ALLSPACE"
        DISPLAY "FAIL: want [ALLSPACE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. END-IF. STOP RUN.

