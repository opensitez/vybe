*> vybe-test: cobol/cobol/func_trim
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. FTRIM.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(20) VALUE "  Hello  ".
01 WS-OUT  PIC X(20).
PROCEDURE DIVISION.
    MOVE FUNCTION TRIM(WS-TEXT) TO WS-OUT.
    DISPLAY WS-OUT.
    STOP RUN.

