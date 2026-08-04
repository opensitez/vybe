*> vybe-test: cobol/category_string_functions/test_str_fn_test_date_yyyymmdd_invalid_day
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION TEST-DATE-YYYYMMDD(20230132).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE TEST-DATE-YYYYMMDD(20230132) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

