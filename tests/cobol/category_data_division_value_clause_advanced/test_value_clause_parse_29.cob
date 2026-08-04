*> vybe-test: cobol/category_data_division_value_clause_advanced/test_value_clause_parse_29
*> origin: languages/cobol/tests/cobol/test_category_data_division_value_clause_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 NEG PIC S9(4) VALUE -12. PROCEDURE DIVISION. IF NEG < 0 DISPLAY 'NEG'.
    MOVE SPACES TO WS-VYBE-L
    STRING 'NEG' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NEG"
        DISPLAY "FAIL: want [NEG] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. END-IF. STOP RUN.

