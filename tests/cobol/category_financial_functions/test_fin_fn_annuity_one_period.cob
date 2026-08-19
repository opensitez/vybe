*> vybe-test: cobol/category_financial_functions/test_fin_fn_annuity_one_period
*> origin: languages/cobol/tests/cobol/test_category_financial_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION ANNUITY(0.
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION ANNUITY(0 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1.10"
        DISPLAY "FAIL: want [1.10] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.1 1). STOP RUN.

