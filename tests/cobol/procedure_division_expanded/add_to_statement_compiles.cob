*> vybe-test: cobol/procedure_division_expanded/add_to_statement_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(3) VALUE 2.
01 WS-B PIC 9(3) VALUE 3.
PROCEDURE DIVISION.
    ADD WS-A TO WS-B.
    STOP RUN.

