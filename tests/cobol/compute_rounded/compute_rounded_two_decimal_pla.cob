*> vybe-test: cobol/compute_rounded/compute_rounded_two_decimal_places
*> origin: languages/cobol/tests/cobol/test_compute_rounded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC 9(3)V99 VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE R ROUNDED = 10 / 3.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00333"
        DISPLAY "FAIL: want [00333] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

