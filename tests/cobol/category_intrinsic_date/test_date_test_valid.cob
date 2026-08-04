*> vybe-test: cobol/category_intrinsic_date/test_date_test_valid
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_date.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION TEST-DATE-YYYYMMDD(20230101).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE TEST-DATE-YYYYMMDD(20230101) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0"
        DISPLAY "FAIL: want [0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

