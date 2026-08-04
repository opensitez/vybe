*> vybe-test: cobol/procedure_division_statement_forms/goto_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    GO TO L1.
L1.
    STOP RUN.

