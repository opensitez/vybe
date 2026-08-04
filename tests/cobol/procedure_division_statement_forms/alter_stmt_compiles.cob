*> vybe-test: cobol/procedure_division_statement_forms/alter_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    ALTER L1 TO PROCEED TO L2.
L1. DISPLAY "A".
L2. STOP RUN.

