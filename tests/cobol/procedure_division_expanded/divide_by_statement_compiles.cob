*> vybe-test: cobol/procedure_division_expanded/divide_by_statement_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 9.
01 WS-B PIC 9(3) VALUE 3.
PROCEDURE DIVISION.
    DIVIDE WS-A BY WS-B.
    STOP RUN.

