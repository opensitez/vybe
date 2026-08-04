*> vybe-test: cobol/category_intrinsic_statistical/test_stat_parse_17
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_statistical.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEDIAN(9 1 7 3).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MEDIAN(9 DELIMITED SIZE 1 DELIMITED SIZE 7 DELIMITED SIZE 3) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "5"
        DISPLAY "FAIL: want [5] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

