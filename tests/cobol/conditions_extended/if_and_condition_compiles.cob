*> vybe-test: cobol/conditions_extended/if_and_condition_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 5.
01 WS-B PIC 9(3) VALUE 7.
PROCEDURE DIVISION.
    IF WS-A > 0 AND WS-B > 0
        DISPLAY "A"
    END-IF.
    STOP RUN.

