*> vybe-test: cobol/enterprise/pic_signed
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BAL PIC S9(8)V99 VALUE -500.00.
PROCEDURE DIVISION.
    DISPLAY WS-BAL.
    STOP RUN.

