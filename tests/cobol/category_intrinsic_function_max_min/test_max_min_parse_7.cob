*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_7
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION MAX(1 2 3 4 5).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MAX(1, 2, 3, 4, 5) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

