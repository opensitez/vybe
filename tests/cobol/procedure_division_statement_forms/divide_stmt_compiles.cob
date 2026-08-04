*> vybe-test: cobol/procedure_division_statement_forms/divide_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC 9 VALUE 8.
01 B PIC 9 VALUE 2.
PROCEDURE DIVISION.
    DIVIDE A BY B.
    STOP RUN.

