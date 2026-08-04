*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_alphanumeric
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MAX('APPLE' 'ZEBRA' 'BANANA').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MAX('APPLE' DELIMITED SIZE 'ZEBRA' DELIMITED SIZE 'BANANA') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "ZEBRA"
        DISPLAY "FAIL: want [ZEBRA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

