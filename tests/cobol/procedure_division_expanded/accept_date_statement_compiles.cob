*> vybe-test: cobol/procedure_division_expanded/accept_date_statement_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC X(8).
PROCEDURE DIVISION.
    ACCEPT WS-DATE FROM DATE.
    STOP RUN.

