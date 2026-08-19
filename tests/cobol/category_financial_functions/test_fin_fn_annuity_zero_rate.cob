*> vybe-test: cobol/category_financial_functions/test_fin_fn_annuity_zero_rate
*> origin: languages/cobol/tests/cobol/test_category_financial_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ANNUITY(0 2).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION ANNUITY(0, 2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0.50"
        DISPLAY "FAIL: want [0.50] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

