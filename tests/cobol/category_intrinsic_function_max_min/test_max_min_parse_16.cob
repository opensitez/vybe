*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_16
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIN(10 20).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MIN(10, 20) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "10"
        DISPLAY "FAIL: want [10] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

