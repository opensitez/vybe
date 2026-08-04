*> vybe-test: cobol/procedure_division_expanded/move_field_to_field_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(5) VALUE "HI".
01 WS-B PIC X(5).
PROCEDURE DIVISION.
    MOVE WS-A TO WS-B.
    STOP RUN.

