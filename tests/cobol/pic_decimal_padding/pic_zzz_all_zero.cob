*> vybe-test: cobol/pic_decimal_padding/pic_zzz_all_zero
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC Z(5) VALUE 0.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

