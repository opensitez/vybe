*> vybe-test: cobol/cobol/search_table
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SRCH.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ENTRY OCCURS 10 TIMES.
      10 WS-KEY   PIC 9(3).
      10 WS-VALUE PIC X(10).
01 WS-IDX PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    DISPLAY "Search test".
    STOP RUN.

