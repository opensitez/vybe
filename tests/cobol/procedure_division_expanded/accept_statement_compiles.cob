*> vybe-test: cobol/procedure_division_expanded/accept_statement_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(10).
PROCEDURE DIVISION.
    ACCEPT WS-A.
    STOP RUN.

