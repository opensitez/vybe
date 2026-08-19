*> vybe-test: cobol/category_intrinsic_function_upper_case/test_upper_case_literal
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_upper_case.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION UPPER-CASE('abc').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION UPPER-CASE('abc') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ABC"
        DISPLAY "FAIL: want [ABC] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

