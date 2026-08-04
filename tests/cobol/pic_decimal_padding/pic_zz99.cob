*> vybe-test: cobol/pic_decimal_padding/pic_zz99
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC ZZ99 VALUE 42.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

