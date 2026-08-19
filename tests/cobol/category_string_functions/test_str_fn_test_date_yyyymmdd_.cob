*> vybe-test: cobol/category_string_functions/test_str_fn_test_date_yyyymmdd_invalid_month
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION TEST-DATE-YYYYMMDD(20231301).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION TEST-DATE-YYYYMMDD(20231301) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

