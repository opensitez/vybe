*> vybe-test: cobol/cobol2023_nested/global_data_item
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SHARED PIC X(20) GLOBAL.
01 WS-LOCAL PIC X(20).
PROCEDURE DIVISION.
    MOVE "Global value" TO WS-SHARED.
    DISPLAY WS-SHARED.
    STOP RUN.

