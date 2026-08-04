*> vybe-test: cobol/enterprise/global_item
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SHARED PIC X(50) GLOBAL.
PROCEDURE DIVISION.
    MOVE "Hello" TO WS-SHARED.
    DISPLAY WS-SHARED.
    STOP RUN.

