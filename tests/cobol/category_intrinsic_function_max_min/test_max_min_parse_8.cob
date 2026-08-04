*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_8
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIN(9 8 7 6).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MIN(9 DELIMITED SIZE 8 DELIMITED SIZE 7 DELIMITED SIZE 6) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "6"
        DISPLAY "FAIL: want [6] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

