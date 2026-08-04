*> vybe-test: cobol/category_string_functions/test_str_fn_min
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIN('Z' 'X' 'Y').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MIN('Z' DELIMITED SIZE 'X' DELIMITED SIZE 'Y') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "X"
        DISPLAY "FAIL: want [X] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

