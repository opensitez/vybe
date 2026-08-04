*> vybe-test: cobol/category_intrinsic_date/test_day_test_invalid_day
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_date.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION TEST-DAY-YYYYDDD(2023366).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE TEST-DAY-YYYYDDD(2023366) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "2"
        DISPLAY "FAIL: want [2] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

