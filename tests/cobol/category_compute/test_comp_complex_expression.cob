*> vybe-test: cobol/category_compute/test_comp_complex_expression
*> origin: languages/cobol/tests/cobol/test_category_compute.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC 999. PROCEDURE DIVISION. COMPUTE R = 10 + 5 * 2 ** 3 / 4 - 1. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "019"
        DISPLAY "FAIL: want [019] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

