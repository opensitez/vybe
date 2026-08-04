*> vybe-test: cobol/category_arithmetic_rounding/test_round_div_negative_result
*> origin: languages/cobol/tests/cobol/test_category_arithmetic_rounding.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. DATA DIVISION. WORKING-STORAGE SECTION.
01 WS-VYBE-L PIC X(256). 01 R PIC S99. PROCEDURE DIVISION. DIVIDE 6 INTO -25 GIVING R ROUNDED. IF R IS NEGATIVE DISPLAY "NEG" ELSE DISPLAY "POS" END-IF.
    MOVE SPACES TO WS-VYBE-L
    STRING "NEG" DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "NEG"
        DISPLAY "FAIL: want [NEG] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

