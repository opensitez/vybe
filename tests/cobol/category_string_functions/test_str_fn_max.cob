*> vybe-test: cobol/category_string_functions/test_str_fn_max
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MAX('A' 'C' 'B').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MAX('A' DELIMITED SIZE 'C' DELIMITED SIZE 'B') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "C"
        DISPLAY "FAIL: want [C] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

