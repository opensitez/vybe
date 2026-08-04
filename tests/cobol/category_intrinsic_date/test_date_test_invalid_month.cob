*> vybe-test: cobol/category_intrinsic_date/test_date_test_invalid_month
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_date.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION TEST-DATE-YYYYMMDD(20231301).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE TEST-DATE-YYYYMMDD(20231301) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

