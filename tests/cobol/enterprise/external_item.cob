*> vybe-test: cobol/enterprise/external_item
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-EXT PIC X(50) EXTERNAL.
PROCEDURE DIVISION.
    DISPLAY WS-EXT.
    STOP RUN.

