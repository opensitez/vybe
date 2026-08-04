*> vybe-test: cobol/final_features/occurs_depending_on
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. ODO.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-COUNT PIC 9(3) VALUE 5.
01 WS-TABLE.
   05 WS-ITEM PIC X(10) OCCURS 1 TO 100 TIMES
      DEPENDING ON WS-COUNT.
PROCEDURE DIVISION.
    DISPLAY "ODO Test".
    STOP RUN.

