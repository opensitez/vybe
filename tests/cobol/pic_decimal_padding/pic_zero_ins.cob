*> vybe-test: cobol/pic_decimal_padding/pic_zero_ins
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(3)09(3) VALUE 123456.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

