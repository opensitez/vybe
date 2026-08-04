*> vybe-test: cobol/structs_and_complex_records/redefines_complex_record_compiles
*> origin: languages/cobol/tests/cobol/test_structs_and_complex_records.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BLOCK PIC X(30).
01 WS-BLOCK-VIEW REDEFINES WS-BLOCK.
   05 WS-CODE PIC X(5).
   05 WS-AMOUNT PIC 9(5).
   05 WS-DESC PIC X(20).
PROCEDURE DIVISION.
    MOVE "A1000" TO WS-CODE.
    MOVE 123 TO WS-AMOUNT.
    STOP RUN.

