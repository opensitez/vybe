*> vybe-test: cobol/category_intrinsic_statistical/test_stat_parse_19
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_statistical.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION RANGE(-10 0 20).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE RANGE(-10 DELIMITED SIZE 0 DELIMITED SIZE 20) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "30"
        DISPLAY "FAIL: want [30] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

