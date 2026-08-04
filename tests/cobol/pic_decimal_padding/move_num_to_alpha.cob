*> vybe-test: cobol/pic_decimal_padding/move_num_to_alpha
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NUM PIC 9(5) VALUE 12345.
01 WS-STR PIC X(10).
PROCEDURE DIVISION.
    MOVE WS-NUM TO WS-STR.
    DISPLAY WS-STR.
    STOP RUN.

