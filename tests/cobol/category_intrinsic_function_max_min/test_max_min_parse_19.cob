*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_19
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MAX(1.
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MAX(1 DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3.7"
        DISPLAY "FAIL: want [3.7] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.2 3.7 2.1). STOP RUN.

