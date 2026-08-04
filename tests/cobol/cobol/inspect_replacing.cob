*> vybe-test: cobol/cobol/inspect_replacing
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. INSPR.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(20) VALUE "Hello World".
PROCEDURE DIVISION.
    INSPECT WS-TEXT REPLACING ALL "l" BY "r".
    DISPLAY WS-TEXT.
    STOP RUN.

