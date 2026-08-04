*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_15
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MAX(10 20).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MAX(10 DELIMITED SIZE 20) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "20"
        DISPLAY "FAIL: want [20] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

