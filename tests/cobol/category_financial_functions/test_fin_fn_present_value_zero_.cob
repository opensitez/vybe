*> vybe-test: cobol/category_financial_functions/test_fin_fn_present_value_zero_rate
*> origin: languages/cobol/tests/cobol/test_category_financial_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(0 100 100).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION PRESENT-VALUE(0, 100, 100) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "200.00"
        DISPLAY "FAIL: want [200.00] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

