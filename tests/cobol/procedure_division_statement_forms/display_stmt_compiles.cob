*> vybe-test: cobol/procedure_division_statement_forms/display_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(5) VALUE "A".
PROCEDURE DIVISION.
    DISPLAY A.
    STOP RUN.

