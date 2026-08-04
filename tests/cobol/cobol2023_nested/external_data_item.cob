*> vybe-test: cobol/cobol2023_nested/external_data_item
*> origin: languages/cobol/tests/cobol/test_cobol2023_nested.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-EXT PIC X(20) EXTERNAL.
PROCEDURE DIVISION.
    DISPLAY WS-EXT.
    STOP RUN.

