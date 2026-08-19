*> vybe-test: cobol/category_intrinsic_function_reverse/test_reverse_literal
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_reverse.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION REVERSE('12345').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION REVERSE('12345') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "54321"
        DISPLAY "FAIL: want [54321] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

