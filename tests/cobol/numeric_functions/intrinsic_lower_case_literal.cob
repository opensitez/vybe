*> vybe-test: cobol/numeric_functions/intrinsic_lower_case_literal
*> origin: languages/cobol/tests/cobol/test_numeric_functions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC X(10).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE FUNCTION LOWER-CASE("HELLO") TO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "hello     "
        DISPLAY "FAIL: want [hello     ] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

