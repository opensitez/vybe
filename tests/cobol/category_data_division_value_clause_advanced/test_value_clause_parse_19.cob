*> vybe-test: cobol/category_data_division_value_clause_advanced/test_value_clause_parse_19
*> origin: languages/cobol/tests/cobol/test_category_data_division_value_clause_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X VALUE LOW-VALUE. PROCEDURE DIVISION. IF V = LOW-VALUE DISPLAY 'LOW'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'LOW' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "LOW"
        DISPLAY "FAIL: want [LOW] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. END-IF. STOP RUN.

