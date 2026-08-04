*> vybe-test: cobol/category_compute_rounded_mode/test_comp_rnd_parse_26
*> origin: languages/cobol/tests/cobol/test_category_compute_rounded_mode.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 X PIC 99 VALUE 0. PROCEDURE DIVISION. COMPUTE X ROUNDED MODE IS TOWARD-ZERO = 12.8. DISPLAY X.
    MOVE SPACES TO WS-VYBE-L
    STRING X DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "12"
        DISPLAY "FAIL: want [12] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

