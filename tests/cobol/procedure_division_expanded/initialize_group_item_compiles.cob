*> vybe-test: cobol/procedure_division_expanded/initialize_group_item_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-A PIC X(3) VALUE "AB".
   05 WS-B PIC 9(2) VALUE 10.
PROCEDURE DIVISION.
    INITIALIZE WS-REC.
    STOP RUN.

