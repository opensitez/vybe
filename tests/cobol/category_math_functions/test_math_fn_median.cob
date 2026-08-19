*> vybe-test: cobol/category_math_functions/test_math_fn_median
*> origin: languages/cobol/tests/cobol/test_category_math_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION MEDIAN(1 9 5).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MEDIAN(1, 9, 5) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

