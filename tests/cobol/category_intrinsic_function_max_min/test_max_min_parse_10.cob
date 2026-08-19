*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_10
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION MIN('ZZZ' 'AAA' 'MMM').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MIN('ZZZ', 'AAA', 'MMM') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "AAA"
        DISPLAY "FAIL: want [AAA] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

