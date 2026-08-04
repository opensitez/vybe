*> vybe-test: cobol/pic_clauses/test_pic_implicit_decimal
*> origin: languages/cobol/tests/cobol/test_pic_clauses.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DEC PIC 9(5)V99 VALUE 1234.56.
01 WS-DEC-ONLY PIC V999 VALUE .123.
01 WS-SIGNED-DEC PIC S9(7)V99 VALUE -1234.56.
PROCEDURE DIVISION.

    DISPLAY WS-DEC.
    STOP RUN.

