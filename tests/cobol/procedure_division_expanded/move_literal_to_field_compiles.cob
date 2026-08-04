*> vybe-test: cobol/procedure_division_expanded/move_literal_to_field_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(10).
PROCEDURE DIVISION.
    MOVE "ALICE" TO WS-NAME.
    STOP RUN.

