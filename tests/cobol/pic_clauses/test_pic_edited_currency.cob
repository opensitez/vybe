*> vybe-test: cobol/pic_clauses/test_pic_edited_currency
*> origin: languages/cobol/tests/cobol/test_pic_clauses.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-EDIT1 PIC $9(6) VALUE ZERO.
01 WS-EDIT2 PIC $$$$$9 VALUE ZERO.
PROCEDURE DIVISION.

    MOVE 123 TO WS-EDIT1.
    MOVE 123 TO WS-EDIT2.
    STOP RUN.

