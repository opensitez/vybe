*> vybe-test: cobol/cobol/func_substitute
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. FSUB.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(30) VALUE "Hello World".
01 WS-OUT  PIC X(30).
PROCEDURE DIVISION.
    MOVE FUNCTION SUBSTITUTE(WS-TEXT "World" "COBOL")
         TO WS-OUT.
    DISPLAY WS-OUT.
    STOP RUN.

