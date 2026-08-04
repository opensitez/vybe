*> vybe-test: cobol/category_intrinsic_date/test_date_integer_min
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_date.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION INTEGER-OF-DATE(16010101).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE INTEGER-OF-DATE(16010101) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "1"
        DISPLAY "FAIL: want [1] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

