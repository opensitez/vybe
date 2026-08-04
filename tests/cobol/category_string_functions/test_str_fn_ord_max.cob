*> vybe-test: cobol/category_string_functions/test_str_fn_ord_max
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION ORD-MAX('A' 'C' 'B').
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE ORD-MAX('A' DELIMITED SIZE 'C' DELIMITED SIZE 'B') DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

