*> vybe-test: cobol/category_data_division_value_clause_advanced/test_value_clause_parse_28
*> origin: languages/cobol/tests/cobol/test_category_data_division_value_clause_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(6) VALUE 'ABC' . 01 L PIC X(3) VALUE ALL 'Z'. PROCEDURE DIVISION. DISPLAY '[' V '|' L ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE V DELIMITED SIZE '|' DELIMITED SIZE L DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[ABC   |ZZZ]"
        DISPLAY "FAIL: want [[ABC   |ZZZ]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

