*> vybe-test: cobol/procedure_division_extended/procedure_division_perform_statement_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 2 TIMES
        DISPLAY "X"
    END-PERFORM.
    STOP RUN.

