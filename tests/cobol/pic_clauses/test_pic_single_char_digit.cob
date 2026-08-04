*> vybe-test: cobol/pic_clauses/test_pic_single_char_digit
*> origin: languages/cobol/tests/cobol/test_pic_clauses.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CHAR PIC X VALUE "A".
01 WS-DIGIT PIC 9 VALUE 7.
PROCEDURE DIVISION.

    DISPLAY WS-CHAR.
    DISPLAY WS-DIGIT.
    STOP RUN.

