*> vybe-test: cobol/category_compute_rounded_mode/test_comp_rnd_parse_24
*> origin: languages/cobol/tests/cobol/test_category_compute_rounded_mode.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC S99 VALUE 0.01. 01 S PIC 99 VALUE 0. PROCEDURE DIVISION. COMPUTE R = S + 0.5. COMPUTE S ROUNDED MODE IS NEAREST-EVEN = R + 1.5. DISPLAY S.
    MOVE SPACES TO WS-VYBE-L
    STRING S DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "02"
        DISPLAY "FAIL: want [02] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

