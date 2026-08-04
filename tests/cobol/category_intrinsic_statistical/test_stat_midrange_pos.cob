*> vybe-test: cobol/category_intrinsic_statistical/test_stat_midrange_pos
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_statistical.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MIDRANGE(2 10).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MIDRANGE(2 DELIMITED SIZE 10) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "6"
        DISPLAY "FAIL: want [6] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

