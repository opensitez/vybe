*> vybe-test: cobol/pic_decimal_padding/move_alpha_truncate
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC X(20) VALUE "Hello World".
01 WS-DST PIC X(5).
PROCEDURE DIVISION.
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
    STOP RUN.

