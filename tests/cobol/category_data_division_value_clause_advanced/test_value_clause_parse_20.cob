*> vybe-test: cobol/category_data_division_value_clause_advanced/test_value_clause_parse_20
*> origin: languages/cobol/tests/cobol/test_category_data_division_value_clause_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 77 V-77 PIC 9(3) VALUE 9. PROCEDURE DIVISION. DISPLAY V-77.
    MOVE SPACES TO WS-VYBE-L
    STRING V-77 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "009"
        DISPLAY "FAIL: want [009] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

