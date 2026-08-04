*> vybe-test: cobol/procedure_division_statement_forms/perform_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 2 TIMES DISPLAY "L" END-PERFORM.
    STOP RUN.

