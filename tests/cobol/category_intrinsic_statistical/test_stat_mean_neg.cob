*> vybe-test: cobol/category_intrinsic_statistical/test_stat_mean_neg
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_statistical.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEAN(-2 -4 -6 -8).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MEAN(-2 DELIMITED SIZE -4 DELIMITED SIZE -6 DELIMITED SIZE -8) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "-5"
        DISPLAY "FAIL: want [-5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

