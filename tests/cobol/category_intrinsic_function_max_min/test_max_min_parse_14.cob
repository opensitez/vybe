*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_14
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIN(1 1 1).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MIN(1 DELIMITED SIZE 1 DELIMITED SIZE 1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

