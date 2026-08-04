*> vybe-test: cobol/signed_arithmetic/signed_compute_absolute_value_simulation
*> origin: languages/cobol/tests/cobol/test_signed_arithmetic.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC S9(4) VALUE -25.
01 ABS-X PIC 9(4) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    IF X < 0
        COMPUTE ABS-X = -X
    ELSE
        MOVE X TO ABS-X
    END-IF.
    DISPLAY ABS-X.
    MOVE SPACES TO WS-VYBE-L
    STRING ABS-X DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "0025"
        DISPLAY "FAIL: want [0025] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

