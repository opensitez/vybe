*> vybe-test: cobol/category_math_functions/test_math_fn_sum
*> origin: languages/cobol/tests/cobol/test_category_math_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SUM(1 2 3).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE SUM(1 DELIMITED SIZE 2 DELIMITED SIZE 3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "6"
        DISPLAY "FAIL: want [6] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

