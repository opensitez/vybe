*> vybe-test: cobol/cobol/inspect_tallying
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. INSP.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT  PIC X(20) VALUE "Hello World".
01 WS-COUNT PIC 9(3)  VALUE 0.
PROCEDURE DIVISION.
    INSPECT WS-TEXT TALLYING WS-COUNT FOR ALL "l".
    DISPLAY WS-COUNT.
    STOP RUN.

