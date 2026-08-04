*> vybe-test: cobol/pic_decimal_padding/comp3_rounding
*> origin: languages/cobol/tests/cobol/test_pic_decimal_padding.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-A PIC 9(5)V99 COMP-3 VALUE 0.
01 WS-B PIC 9(5)V99 COMP-3 VALUE 10.00.
01 WS-C PIC 9(5)V99 COMP-3 VALUE 3.
PROCEDURE DIVISION.
    COMPUTE WS-A = WS-B / WS-C.
    DISPLAY WS-A.
    STOP RUN.

