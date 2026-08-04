*> vybe-test: cobol/conditions_extended/if_not_condition_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 0.
PROCEDURE DIVISION.
    IF NOT WS-A > 0
        DISPLAY "A"
    END-IF.
    STOP RUN.

