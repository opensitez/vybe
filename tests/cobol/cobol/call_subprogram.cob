*> vybe-test: cobol/cobol/call_subprogram
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. CALLER.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RESULT PIC 9(5).
PROCEDURE DIVISION.
    CALL "SUBPROG" USING WS-RESULT.
    DISPLAY WS-RESULT.
    STOP RUN.

