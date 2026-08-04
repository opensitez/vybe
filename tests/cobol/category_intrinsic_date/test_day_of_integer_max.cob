*> vybe-test: cobol/category_intrinsic_date/test_day_of_integer_max
*> origin: languages/cobol/tests/cobol/test_category_intrinsic_date.rs
IDENTIFICATION DIVISION. PROGRAM-ID. T. PROCEDURE DIVISION. DISPLAY FUNCTION DAY-OF-INTEGER(3067671).
    MOVE SPACES TO WS-VYBE-L
    STRING FUNCTION DELIMITED SIZE DAY-OF-INTEGER(3067671) DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "9999365"
        DISPLAY "FAIL: want [9999365] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF. STOP RUN.

