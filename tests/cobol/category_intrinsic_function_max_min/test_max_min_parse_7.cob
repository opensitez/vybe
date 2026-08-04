*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_7
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MAX(1 2 3 4 5).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MAX(1 DELIMITED SIZE 2 DELIMITED SIZE 3 DELIMITED SIZE 4 DELIMITED SIZE 5) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

