*> vybe-test: cobol/procedure_division_expanded/initialize_data_item_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(5) VALUE "AB".
PROCEDURE DIVISION.
    INITIALIZE WS-NAME.
    STOP RUN.

