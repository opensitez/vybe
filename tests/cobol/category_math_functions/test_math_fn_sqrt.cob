*> vybe-test: cobol/category_math_functions/test_math_fn_sqrt
*> origin: languages/cobol/tests/cobol/test_category_math_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SQRT(16).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE SQRT(16) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "4"
        DISPLAY "FAIL: want [4] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

