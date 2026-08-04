*> vybe-test: cobol/cobol/occurs_table
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. TABLES.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ITEM PIC X(10) OCCURS 5 TIMES.
PROCEDURE DIVISION.
    MOVE "First"  TO WS-ITEM(1).
    MOVE "Second" TO WS-ITEM(2).
    DISPLAY WS-ITEM(1).
    DISPLAY WS-ITEM(2).
    STOP RUN.

