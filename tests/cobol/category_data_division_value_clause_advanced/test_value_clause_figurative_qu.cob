*> vybe-test: cobol/category_data_division_value_clause_advanced/test_value_clause_figurative_quote
*> origin: languages/cobol/tests/cobol/test_category_data_division_value_clause_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(2) VALUE QUOTE. PROCEDURE DIVISION. DISPLAY V.
    MOVE SPACES TO WS-VYBE-L
    STRING V DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = """"
        DISPLAY "FAIL: want [""] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

