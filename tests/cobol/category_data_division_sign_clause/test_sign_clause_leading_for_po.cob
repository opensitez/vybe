*> vybe-test: cobol/category_data_division_sign_clause/test_sign_clause_leading_for_positive_with_no_override
*> origin: languages/cobol/tests/cobol/test_category_data_division_sign_clause.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 V PIC S9(2) SIGN IS LEADING VALUE +12. PROCEDURE DIVISION. DISPLAY V.
    MOVE SPACES TO WS-VYBE-L
    STRING V DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "+12"
        DISPLAY "FAIL: want [+12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

