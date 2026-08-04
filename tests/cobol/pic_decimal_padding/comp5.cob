*> vybe-test: cobol/pic_decimal_padding/comp5
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(9) USAGE BINARY VALUE 255.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

