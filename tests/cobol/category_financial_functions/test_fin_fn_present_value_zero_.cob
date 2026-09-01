*> vybe-test: cobol/category_financial_functions/test_fin_fn_present_value_zero_rate
*> ⛔ EXPECTATION CORRECTED 2026-08-29. This asserted [0.50] / [200.00],
*> which only ever passed because the walker forced every financial
*> intrinsic through a 2-decimal formatter. cobc, asked the same question
*> by an equivalent minimal program (this one's own checker names an
*> undefined item, so cobc cannot judge it directly):
*>     DISPLAY FUNCTION ANNUITY(0 2)             -> 00000000.5
*>     DISPLAY FUNCTION PRESENT-VALUE(0 100 100) -> 000000200
*> An intrinsic returns a value; nothing here supplies a PICTURE, so no
*> trailing zero exists to print.
*> origin: languages/cobol/tests/cobol/test_category_financial_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION PRESENT-VALUE(0 100 100).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION PRESENT-VALUE(0, 100, 100) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "200"
        DISPLAY "FAIL: want [200] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

