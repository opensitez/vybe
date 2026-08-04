*> vybe-test: cobol/parse_debug/parse_subscript_access
*> origin: languages/cobol/tests/cobol/test_parse_debug.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TABLE.
   05 WS-ITEM OCCURS 10 TIMES PIC 9(3).
01 WS-IDX PIC 9(2) VALUE 1.
PROCEDURE DIVISION.
    DISPLAY WS-ITEM(WS-IDX).
    STOP RUN.

