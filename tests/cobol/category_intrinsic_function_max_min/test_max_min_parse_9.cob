*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_9
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION MAX('A' 'Z' 'F').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MAX('A', 'Z', 'F') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Z"
        DISPLAY "FAIL: want [Z] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

