*> vybe-test: cobol/pic_decimal_padding/move_spaces_to_alpha
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC X(20) VALUE "Old data".
PROCEDURE DIVISION.
    MOVE SPACES TO WS-X.
    DISPLAY WS-X.
    STOP RUN.

