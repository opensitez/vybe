*> vybe-test: cobol/category_data_division_value_clause_advanced/test_value_clause_parse_21
*> origin: languages/cobol/tests/cobol/test_category_data_division_value_clause_advanced.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC X(3) VALUE 'ABC'. 77 W PIC 9(4) VALUE 7. PROCEDURE DIVISION. DISPLAY '[' V ']'.
    MOVE SPACES TO WS-VYBE-L
    STRING '[' DELIMITED SIZE V DELIMITED SIZE ']' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "[ABC]"
        DISPLAY "FAIL: want [[ABC]] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. DISPLAY W.
    MOVE SPACES TO WS-VYBE-L
    STRING W DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0007"
        DISPLAY "FAIL: want [0007] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

