*> vybe-test: cobol/category_intrinsic_trigonometric/test_trig_cos_0
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_trigonometric.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION COS(0).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE COS(0) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

