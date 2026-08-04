*> vybe-test: cobol/pic_decimal_padding/comp_binary
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(9) COMP VALUE 12345.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

