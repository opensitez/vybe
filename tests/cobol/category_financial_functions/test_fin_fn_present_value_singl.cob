*> vybe-test: cobol/category_financial_functions/test_fin_fn_present_value_single
*> origin: languages/cobol/tests/cobol/test_category_financial_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(0.
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION PRESENT-VALUE(0 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "100.00"
        DISPLAY "FAIL: want [100.00] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.1 110). STOP RUN.

