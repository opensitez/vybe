*> vybe-test: cobol/category_financial_functions/test_fin_fn_annuity
*> origin: languages/cobol/tests/cobol/test_category_financial_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ANNUITY(0.
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE ANNUITY(0 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0.53"
        DISPLAY "FAIL: want [0.53] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.05 2). STOP RUN.

