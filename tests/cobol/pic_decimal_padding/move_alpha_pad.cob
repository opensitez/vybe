*> vybe-test: cobol/pic_decimal_padding/move_alpha_pad
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC X(20) VALUE "Hello".
PROCEDURE DIVISION.
    DISPLAY WS-A.
    STOP RUN.

