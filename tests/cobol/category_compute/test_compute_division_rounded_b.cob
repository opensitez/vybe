*> vybe-test: cobol/category_compute/test_compute_division_rounded_by_target
*> origin: languages/cobol/tests/cobol/test_category_compute.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC 9. 01 X PIC 9(2) VALUE 10. 01 Y PIC 9(2) VALUE 3. PROCEDURE DIVISION. COMPUTE R ROUNDED = X / Y. DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "3"
        DISPLAY "FAIL: want [3] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

