*> vybe-test: cobol/pic_clauses/test_pic_signed_display
*> origin: languages/cobol/tests/cobol/test_pic_clauses.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

01 WS-POS PIC S9(5) VALUE +123.
01 WS-NEG PIC S9(5) VALUE -123.
    STOP RUN.

