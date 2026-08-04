*> vybe-test: cobol/category_math_functions/test_math_fn_integer_part
*> origin: languages/cobol/tests/cobol/test_category_math_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION INTEGER-PART(2.
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE INTEGER-PART(2 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.5). STOP RUN.

