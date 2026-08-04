*> vybe-test: cobol/structs_and_complex_records/move_corresponding_between_complex_records_compiles
*> origin: languages/cobol/tests/cobol/test_structs_and_complex_records.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC.
   05 WS-NAME PIC X(10) VALUE "ITEM".
   05 WS-COUNT PIC 9(4) VALUE 5.
01 WS-DST.
   05 WS-NAME PIC X(10).
   05 WS-COUNT PIC 9(4).
PROCEDURE DIVISION.
    MOVE CORRESPONDING WS-SRC TO WS-DST.
    STOP RUN.

