*> vybe-test: cobol/pic_decimal_padding/pic_v_decimal
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(5)V99 VALUE 123.45.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

