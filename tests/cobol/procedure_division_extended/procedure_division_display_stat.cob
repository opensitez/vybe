*> vybe-test: cobol/procedure_division_extended/procedure_division_display_statement_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(3) VALUE "ABC".
PROCEDURE DIVISION.
    DISPLAY WS-A.
    STOP RUN.

