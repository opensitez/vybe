*> vybe-test: cobol/category_string_functions/test_str_fn_day_of_integer
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION DAY-OF-INTEGER(1).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE DAY-OF-INTEGER(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1601001"
        DISPLAY "FAIL: want [1601001] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

