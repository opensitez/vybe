*> vybe-test: cobol/pic_clauses/test_pic_long_equivalence
*> origin: languages/cobol/tests/cobol/test_pic_clauses.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X4 PIC XXXX VALUE "ABCD".
01 WS-X4-EQ PIC X(4) VALUE "ABCD".
01 WS-94 PIC 9999 VALUE 1234.
01 WS-94-EQ PIC 9(4) VALUE 1234.
PROCEDURE DIVISION.

    DISPLAY WS-X4.
    DISPLAY WS-94.
    STOP RUN.

