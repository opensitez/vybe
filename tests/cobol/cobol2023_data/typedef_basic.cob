*> vybe-test: cobol/cobol2023_data/typedef_basic
*> origin: languages/cobol/tests/cobol/test_cobol2023_data.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AMOUNT PIC 9(7)V99.
PROCEDURE DIVISION.
    TYPEDEF AMOUNT-TYPE AS PIC 9(7)V99.
    STOP RUN.

