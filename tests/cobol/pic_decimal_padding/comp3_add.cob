*> vybe-test: cobol/pic_decimal_padding/comp3_add
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5)V99 COMP-3 VALUE 100.50.
01 WS-B PIC 9(5)V99 COMP-3 VALUE 200.75.
01 WS-C PIC 9(5)V99 COMP-3 VALUE 0.
PROCEDURE DIVISION.
    ADD WS-A TO WS-B.
    DISPLAY WS-B.
    STOP RUN.

