*> vybe-test: cobol/category_compute_rounded_mode/test_comp_rnd_parse_30
*> origin: languages/cobol/tests/cobol/test_category_compute_rounded_mode.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 W PIC 99 VALUE 0. 01 Y PIC 99 VALUE 0. PROCEDURE DIVISION. COMPUTE W = 4.4 + 5.5. COMPUTE Y ROUNDED MODE IS NEAREST-EVEN = W / 1.5. DISPLAY Y.
    MOVE SPACES TO WS-VYBE-L
    STRING Y DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "06"
        DISPLAY "FAIL: want [06] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

