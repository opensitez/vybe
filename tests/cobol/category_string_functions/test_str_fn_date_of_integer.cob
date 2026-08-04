*> vybe-test: cobol/category_string_functions/test_str_fn_date_of_integer
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION DATE-OF-INTEGER(1).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE DATE-OF-INTEGER(1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "16010101"
        DISPLAY "FAIL: want [16010101] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

