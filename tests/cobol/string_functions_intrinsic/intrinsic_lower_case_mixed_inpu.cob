*> vybe-test: cobol/string_functions_intrinsic/intrinsic_lower_case_mixed_input
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "WORLD".
01 R PIC X(5).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE FUNCTION LOWER-CASE(S) TO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "world"
        DISPLAY "FAIL: want [world] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

