*> vybe-test: cobol/category_intrinsic_statistical/test_stat_parse_16
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_statistical.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION MEAN(-1 0 1).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE MEAN(-1 DELIMITED SIZE 0 DELIMITED SIZE 1) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0"
        DISPLAY "FAIL: want [0] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

