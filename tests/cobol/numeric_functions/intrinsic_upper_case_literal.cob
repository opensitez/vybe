*> vybe-test: cobol/numeric_functions/intrinsic_upper_case_literal
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC X(10).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE FUNCTION UPPER-CASE("hello") TO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "HELLO     "
        DISPLAY "FAIL: want [HELLO     ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

