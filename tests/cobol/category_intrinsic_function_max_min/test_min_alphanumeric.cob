*> vybe-test: cobol/category_intrinsic_function_max_min/test_min_alphanumeric
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION MIN('APPLE' 'ZEBRA' 'BANANA').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MIN('APPLE', 'ZEBRA', 'BANANA') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "APPLE"
        DISPLAY "FAIL: want [APPLE] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

