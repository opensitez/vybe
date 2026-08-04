*> vybe-test: cobol/conditions_extended/if_alphanumeric_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(3) VALUE "A1B".
PROCEDURE DIVISION.
    IF WS-A IS ALPHANUMERIC
        DISPLAY "ALNUM"
    END-IF.
    STOP RUN.

