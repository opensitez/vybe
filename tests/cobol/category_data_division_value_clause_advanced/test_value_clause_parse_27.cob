*> vybe-test: cobol/category_data_division_value_clause_advanced/test_value_clause_parse_27
*> origin: languages/cobol/tests/cobol/test_category_data_division_value_clause_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC 9(1) VALUE 8. 88 IS-EIGHT VALUE 8. PROCEDURE DIVISION. IF IS-EIGHT DISPLAY 'EIGHT'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'EIGHT' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "EIGHT"
        DISPLAY "FAIL: want [EIGHT] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. END-IF. STOP RUN.

