*> vybe-test: cobol/category_intrinsic_statistical/test_stat_sum_large
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_statistical.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION SUM(9999 9999).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE SUM(9999 DELIMITED SIZE 9999) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "19998"
        DISPLAY "FAIL: want [19998] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

