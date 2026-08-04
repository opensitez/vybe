*> vybe-test: cobol/pic_clauses/test_pic_alphabetic
*> origin: languages/cobol/tests/cobol/test_pic_clauses.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC A(10) VALUE "HELLO".
01 WS-B PIC A VALUE "X".
PROCEDURE DIVISION.

    DISPLAY WS-A.
    DISPLAY WS-B.
    STOP RUN.

