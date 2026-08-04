*> vybe-test: cobol/category_string_functions/test_str_fn_test_date_yyyymmdd
*> origin: languages/cobol/tests/cobol/test_category_string_functions.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION TEST-DATE-YYYYMMDD(20230101).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE TEST-DATE-YYYYMMDD(20230101) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0"
        DISPLAY "FAIL: want [0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

