*> vybe-test: cobol/pic_clauses/test_pic_national_characters
*> origin: languages/cobol/tests/cobol/test_pic_clauses.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAT PIC N(10) VALUE "HELLO".
PROCEDURE DIVISION.

    DISPLAY WS-NAT.
    STOP RUN.

