*> vybe-test: cobol/conditions_extended/if_numeric_zero_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC S9(3) VALUE 0.
PROCEDURE DIVISION.
    IF WS-A IS ZERO
        DISPLAY "ZERO"
    END-IF.
    STOP RUN.

