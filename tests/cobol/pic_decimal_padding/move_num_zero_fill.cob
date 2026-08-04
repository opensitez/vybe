*> vybe-test: cobol/pic_decimal_padding/move_num_zero_fill
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-SRC PIC 9(3) VALUE 42.
01 WS-DST PIC 9(5).
PROCEDURE DIVISION.
    MOVE WS-SRC TO WS-DST.
    DISPLAY WS-DST.
    STOP RUN.

