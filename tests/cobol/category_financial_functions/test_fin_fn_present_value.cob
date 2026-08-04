*> vybe-test: cobol/category_financial_functions/test_fin_fn_present_value
*> origin: languages/cobol/tests/cobol/test_category_financial_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(0.
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE PRESENT-VALUE(0 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "185.94"
        DISPLAY "FAIL: want [185.94] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.05 100 100). STOP RUN.

