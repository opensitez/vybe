*> vybe-test: cobol/pic_decimal_padding/pic_dollar
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AMT PIC $9(5).99 VALUE 1234.56.
PROCEDURE DIVISION.
    DISPLAY WS-AMT.
    STOP RUN.

