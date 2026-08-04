*> vybe-test: cobol/pic_decimal_padding/move_zeros_to_num
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9(5) VALUE 999.
PROCEDURE DIVISION.
    MOVE ZEROS TO WS-X.
    DISPLAY WS-X.
    STOP RUN.

