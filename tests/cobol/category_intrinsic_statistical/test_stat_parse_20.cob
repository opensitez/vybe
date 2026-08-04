*> vybe-test: cobol/category_intrinsic_statistical/test_stat_parse_20
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_statistical.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. IF FUNCTION MEAN(10 20) = 15 DISPLAY 'Y' END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING 'Y' DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "Y"
        DISPLAY "FAIL: want [Y] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

