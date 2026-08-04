*> vybe-test: cobol/string_functions_intrinsic/intrinsic_reverse_palindrome
*> origin: languages/cobol/tests/cobol/test_string_functions_intrinsic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 S PIC X(5) VALUE "RADAR".
01 R PIC X(5).
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    MOVE FUNCTION REVERSE(S) TO R.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "RADAR"
        DISPLAY "FAIL: want [RADAR] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

