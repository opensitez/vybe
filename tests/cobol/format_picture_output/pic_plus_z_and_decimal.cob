*> vybe-test: cobol/format_picture_output/pic_plus_z_and_decimal
*> origin: languages/cobol/tests/cobol/test_format_picture_output.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC +Z(4).99 VALUE -12.5.
PROCEDURE DIVISION.
    DISPLAY N.
    STOP RUN.

