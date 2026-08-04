*> vybe-test: cobol/category_math_functions/test_math_fn_exp10
*> origin: languages/cobol/tests/cobol/test_category_math_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION EXP10(2).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE EXP10(2) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "100"
        DISPLAY "FAIL: want [100] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

