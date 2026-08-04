*> vybe-test: cobol/cobol2023_data/data_item_indexed_table
*> origin: languages/cobol/tests/cobol/test_cobol2023_data.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ROW OCCURS 10 TIMES
      INDEXED BY WS-IDX.
      10 WS-COL PIC X(5).
PROCEDURE DIVISION.
    DISPLAY WS-COL(1).
    STOP RUN.

