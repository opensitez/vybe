*> vybe-test: cobol/final_features/national_usage
*> origin: languages/cobol/tests/cobol/test_final_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. NATUSAGE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-TEXT PIC X(50) USAGE NATIONAL.
PROCEDURE DIVISION.
    MOVE "Hello World" TO WS-TEXT.
    DISPLAY WS-TEXT.
    STOP RUN.

