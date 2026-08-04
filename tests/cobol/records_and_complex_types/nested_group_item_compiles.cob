*> vybe-test: cobol/records_and_complex_types/nested_group_item_compiles
*> origin: languages/cobol/tests/cobol/test_records_and_complex_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-NAME PIC X(10).
   05 WS-AGE PIC 9(2).
PROCEDURE DIVISION.
    MOVE "ALICE" TO WS-NAME.
    MOVE 30 TO WS-AGE.
    STOP RUN.

