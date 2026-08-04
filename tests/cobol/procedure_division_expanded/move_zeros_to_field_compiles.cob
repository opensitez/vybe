*> vybe-test: cobol/procedure_division_expanded/move_zeros_to_field_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC 9(3).
PROCEDURE DIVISION.
    MOVE ZEROS TO WS-NUM.
    STOP RUN.

