*> vybe-test: cobol/cobol2023_data/data_item_late_binding
*> origin: languages/cobol/tests/cobol/test_cobol2023_data.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DYNAMIC.
   05 WS-NAME PIC X(30) VALUE "Default".
   05 WS-AGE PIC 9(3) VALUE 0.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 100 TIMES.
      10 WS-KEY PIC X(10).
      10 WS-VAL PIC 9(5).
PROCEDURE DIVISION.
    DISPLAY WS-NAME.
    STOP RUN.

