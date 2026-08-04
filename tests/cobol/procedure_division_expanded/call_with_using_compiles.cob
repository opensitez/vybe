*> vybe-test: cobol/procedure_division_expanded/call_with_using_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(5).
PROCEDURE DIVISION.
    CALL "SUBPROG" USING WS-A.
    STOP RUN.

