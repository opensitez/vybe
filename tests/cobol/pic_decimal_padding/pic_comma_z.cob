*> vybe-test: cobol/pic_decimal_padding/pic_comma_z
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC ZZZ,ZZ9 VALUE 12345.
PROCEDURE DIVISION.
    DISPLAY WS-X.
    STOP RUN.

