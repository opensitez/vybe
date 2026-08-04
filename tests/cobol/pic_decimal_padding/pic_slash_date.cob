*> vybe-test: cobol/pic_decimal_padding/pic_slash_date
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC 99/99/9999 VALUE 12252024.
PROCEDURE DIVISION.
    DISPLAY WS-DATE.
    STOP RUN.

