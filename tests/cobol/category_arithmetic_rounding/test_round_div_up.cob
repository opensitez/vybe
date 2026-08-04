*> vybe-test: cobol/category_arithmetic_rounding/test_round_div_up
*> origin: languages/cobol/tests/cobol/test_category_arithmetic_rounding.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC 99. PROCEDURE DIVISION. DIVIDE 3 INTO 20 GIVING R ROUNDED. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "07"
        DISPLAY "FAIL: want [07] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

