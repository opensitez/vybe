*> vybe-test: cobol/format_picture_output/pic_plus_with_decimal_compiles
*> origin: languages/cobol/tests/cobol/test_format_picture_output.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC +9(5).99 VALUE -123.45.
PROCEDURE DIVISION.
    DISPLAY N.
    STOP RUN.

