*> vybe-test: cobol/procedure_division_statement_forms/accept_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
*> vybe-test-mode: compile
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(10).
PROCEDURE DIVISION.
    ACCEPT A.
    STOP RUN.

