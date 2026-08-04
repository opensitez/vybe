*> vybe-test: cobol/procedure_division_statement_forms/set_true_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9.
   88 ONN VALUE 1.
PROCEDURE DIVISION.
    SET ONN TO TRUE.
    STOP RUN.

