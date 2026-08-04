*> vybe-test: cobol/cobol/pic_alpha
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. PICALPHA.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(30) VALUE "Hello World".
PROCEDURE DIVISION.
    DISPLAY WS-TEXT.
    STOP RUN.

