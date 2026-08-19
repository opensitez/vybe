*> vybe-test: cobol/category_intrinsic_function_length/test_length_literal
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_function_length.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION LENGTH('HELLO').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION LENGTH('HELLO') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

