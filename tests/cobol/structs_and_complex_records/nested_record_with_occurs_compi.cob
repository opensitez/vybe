*> vybe-test: cobol/structs_and_complex_records/nested_record_with_occurs_compiles
*> origin: languages/cobol/tests/cobol/test_structs_and_complex_records.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ORDER.
   05 WS-ID PIC 9(6).
   05 WS-LINES OCCURS 3 TIMES.
      10 WS-SKU PIC X(10).
      10 WS-QTY PIC 9(4).
PROCEDURE DIVISION.
    MOVE 1 TO WS-ID.
    MOVE "SKU1" TO WS-SKU(1).
    MOVE 2 TO WS-QTY(1).
    STOP RUN.

