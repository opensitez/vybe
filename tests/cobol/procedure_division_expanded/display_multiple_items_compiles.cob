*> vybe-test: cobol/procedure_division_expanded/display_multiple_items_compiles
*> origin: languages/cobol/tests/cobol/test_procedure_division_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(5) VALUE "A".
01 WS-B PIC 9(2) VALUE 10.
PROCEDURE DIVISION.
    DISPLAY WS-A WS-B.
    STOP RUN.

