*> vybe-test: cobol/conditions_extended/if_alphabetic_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(3) VALUE "ABC".
PROCEDURE DIVISION.
    IF WS-A IS ALPHABETIC
        DISPLAY "ALPHA"
    END-IF.
    STOP RUN.

