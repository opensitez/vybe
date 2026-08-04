*> vybe-test: cobol/conditions_extended/if_less_than_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 3.
01 WS-B PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    IF WS-A < WS-B
        DISPLAY "A"
    END-IF.
    STOP RUN.

