*> vybe-test: cobol/conditions_extended/if_numeric_positive_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC S9(3) VALUE 5.
PROCEDURE DIVISION.
    IF WS-A IS POSITIVE
        DISPLAY "POS"
    END-IF.
    STOP RUN.

