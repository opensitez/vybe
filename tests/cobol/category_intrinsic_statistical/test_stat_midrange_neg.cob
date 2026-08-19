*> vybe-test: cobol/category_intrinsic_statistical/test_stat_midrange_neg
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_statistical.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION. DISPLAY FUNCTION MIDRANGE(-2 -10).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION MIDRANGE(-2, -10) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "����"
        DISPLAY "FAIL: want [����] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

