*> vybe-test: cobol/category_math_functions/test_math_fn_median_even
*> origin: languages/cobol/tests/cobol/test_category_math_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEDIAN(1 5 9 13).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MEDIAN(1 DELIMITED SIZE 5 DELIMITED SIZE 9 DELIMITED SIZE 13) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "7"
        DISPLAY "FAIL: want [7] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

