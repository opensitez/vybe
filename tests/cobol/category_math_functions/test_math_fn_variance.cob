*> vybe-test: cobol/category_math_functions/test_math_fn_variance
*> origin: languages/cobol/tests/cobol/test_category_math_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION VARIANCE(2 4 6).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE VARIANCE(2 DELIMITED SIZE 4 DELIMITED SIZE 6) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

