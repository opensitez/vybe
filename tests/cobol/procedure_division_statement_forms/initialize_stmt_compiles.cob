*> vybe-test: cobol/procedure_division_statement_forms/initialize_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 G.
   05 A PIC X(3) VALUE "A".
PROCEDURE DIVISION.
    INITIALIZE G.
    STOP RUN.

