*> vybe-test: cobol/data_division_extended/data_division_occurs_with_index_compiles
*> origin: languages/cobol/tests/cobol/test_data_division_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TBL.
   05 WS-ITEM PIC 9(2) OCCURS 2 TIMES INDEXED BY WS-IDX.
PROCEDURE DIVISION.
    DISPLAY WS-ITEM(1).
    STOP RUN.

