*> vybe-test: cobol/conditions_extended/perform_inline_condition_compiles
*> origin: languages/cobol/tests/cobol/test_conditions_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 2 TIMES
        CONTINUE
    END-PERFORM.
    STOP RUN.

