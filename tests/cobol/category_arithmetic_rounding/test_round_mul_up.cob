*> vybe-test: cobol/category_arithmetic_rounding/test_round_mul_up
*> origin: languages/cobol/tests/cobol/test_category_arithmetic_rounding.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC 99. PROCEDURE DIVISION. MULTIPLY 2.5 BY 3.0 GIVING R ROUNDED. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "08"
        DISPLAY "FAIL: want [08] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

