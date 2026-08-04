*> vybe-test: cobol/pic_decimal_padding/pic_minus_trail
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(5)- VALUE -100.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

