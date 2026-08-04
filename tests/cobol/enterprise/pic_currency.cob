*> vybe-test: cobol/enterprise/pic_currency
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PRICE PIC 9(5)V99 VALUE 99.99.
PROCEDURE DIVISION.
    DISPLAY WS-PRICE.
    STOP RUN.

