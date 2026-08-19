*> vybe-test: cobol/category_intrinsic_function_lower_case/test_lower_case_literal
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_lower_case.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION LOWER-CASE('ABC').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION LOWER-CASE('ABC') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "abc"
        DISPLAY "FAIL: want [abc] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

