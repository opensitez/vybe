*> vybe-test: cobol/cobol2023_data/data_item_occurs_depending
*> origin: languages/cobol/tests/cobol/test_cobol2023_data.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 9(3) VALUE 5.
01 WS-TABLE.
   05 WS-ITEM OCCURS 1 TO 100 TIMES
      DEPENDING ON WS-COUNT PIC X(10).
PROCEDURE DIVISION.
    DISPLAY WS-COUNT.
    STOP RUN.

