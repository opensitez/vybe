*> vybe-test: cobol/enterprise/pic_z_suppress
*> origin: languages/cobol/tests/cobol/test_enterprise.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AMT PIC 9(8)V99 VALUE 1234.56.
PROCEDURE DIVISION.
    DISPLAY WS-AMT.
    STOP RUN.

