*> vybe-test: cobol/category_math_functions/test_math_fn_integer_part
*> origin: languages/cobol/tests/cobol/test_category_math_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION INTEGER-PART(2.5).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION INTEGER-PART(2.5) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

