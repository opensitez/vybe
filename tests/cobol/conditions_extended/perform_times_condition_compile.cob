*> vybe-test: cobol/conditions_extended/perform_times_condition_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 4 TIMES
        DISPLAY "LOOP"
    END-PERFORM.
    STOP RUN.

