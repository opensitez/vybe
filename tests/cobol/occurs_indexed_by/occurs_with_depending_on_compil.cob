*> vybe-test: cobol/occurs_indexed_by/occurs_with_depending_on_compiles
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 MAX-ITEMS PIC 9(3) VALUE 10.
01 T.
   05 E PIC X(5) OCCURS 1 TO 50 TIMES DEPENDING ON MAX-ITEMS INDEXED BY IX.
PROCEDURE DIVISION.
    SET IX TO 1.
    MOVE "FIRST" TO E(IX).
    STOP RUN.

