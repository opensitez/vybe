*> vybe-test: cobol/category_compute_rounded_mode/test_comp_rnd_toward_zero
*> origin: languages/cobol/tests/cobol/test_category_compute_rounded_mode.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC 99. PROCEDURE DIVISION. COMPUTE R ROUNDED MODE IS TOWARD-ZERO = 10.9. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "10"
        DISPLAY "FAIL: want [10] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

