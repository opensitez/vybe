*> vybe-test: cobol/procedure_division_statement_forms/evaluate_stmt_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_statement_forms.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9 VALUE 1.
PROCEDURE DIVISION.
    EVALUATE X WHEN 1 DISPLAY "A" WHEN OTHER DISPLAY "B" END-EVALUATE.
    STOP RUN.

