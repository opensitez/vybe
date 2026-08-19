*> vybe-test: cobol/category_compute/test_compute_unary_and_negative_mix
*> origin: languages/cobol/tests/cobol/test_category_compute.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC S999. 01 A PIC S9 VALUE -5. 01 B PIC S9 VALUE +3. PROCEDURE DIVISION. COMPUTE R = -A + B * 2. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "011"
        DISPLAY "FAIL: want [011] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

