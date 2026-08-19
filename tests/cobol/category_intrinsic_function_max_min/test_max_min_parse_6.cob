*> vybe-test: cobol/category_intrinsic_function_max_min/test_max_min_parse_6
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_max_min.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 A PIC S9 VALUE -10. 01 B PIC S9 VALUE 5. 01 C PIC S9 VALUE 3. PROCEDURE DIVISION. DISPLAY FUNCTION MIN(A B C).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MIN(A, B, C) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "p"
        DISPLAY "FAIL: want [p] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

