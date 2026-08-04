*> vybe-test: cobol/occurs_indexed_by/occurs_indexed_by_no_redefine
*> origin: languages/cobol/tests/cobol/test_occurs_indexed_by.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 DATA-TABLE.
   05 DT-ENTRY PIC X(10) OCCURS 20 TIMES INDEXED BY DT-IDX.
PROCEDURE DIVISION.
    SET DT-IDX TO 1.
    MOVE "FIRST" TO DT-ENTRY(DT-IDX).
    STOP RUN.

