*> vybe-test: cobol/procedure_division_extended/procedure_division_multiply_statement_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    MULTIPLY 2 BY WS-A.
    STOP RUN.

