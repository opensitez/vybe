*> vybe-test: cobol/category_compute/test_compute_with_precedence
*> origin: languages/cobol/tests/cobol/test_category_compute.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC 999. PROCEDURE DIVISION. COMPUTE R = 2 + 3 * 4 - (8 / 2). DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "010"
        DISPLAY "FAIL: want [010] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

