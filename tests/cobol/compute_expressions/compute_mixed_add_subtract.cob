*> vybe-test: cobol/compute_expressions/compute_mixed_add_subtract
*> origin: languages/cobol/tests/cobol/test_compute_expressions.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 R PIC S9(5) VALUE 0.
01 WS-VYBE-L PIC X(256).
PROCEDURE DIVISION.
    COMPUTE R = 100 - 40 + 15 - 5.
    DISPLAY R.
    MOVE SPACES TO WS-VYBE-L
    STRING R DELIMITED SIZE INTO WS-VYBE-L
    IF WS-VYBE-L NOT = "00070"
        DISPLAY "FAIL: want [00070] got [" WS-VYBE-L "]"
        MOVE 1 TO RETURN-CODE
        RAISE EXCEPTION EC-PROGRAM
    END-IF.
    STOP RUN.

