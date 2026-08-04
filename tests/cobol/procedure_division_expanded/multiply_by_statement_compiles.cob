*> vybe-test: cobol/procedure_division_expanded/multiply_by_statement_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 4.
01 WS-B PIC 9(3) VALUE 5.
PROCEDURE DIVISION.
    MULTIPLY WS-A BY WS-B.
    STOP RUN.

