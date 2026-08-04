*> vybe-test: cobol/records_and_complex_types/redefines_group_item_compiles
*> origin: languages/cobol/tests/cobol/test_records_and_complex_types.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BUFFER PIC X(20).
01 WS-FIELD REDEFINES WS-BUFFER.
   05 WS-CHAR PIC X(1) OCCURS 20 TIMES.
PROCEDURE DIVISION.
    MOVE "A" TO WS-CHAR(1).
    STOP RUN.

