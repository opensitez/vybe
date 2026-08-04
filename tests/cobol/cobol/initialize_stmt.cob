*> vybe-test: cobol/cobol/initialize_stmt
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. INIT.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC.
   05 WS-NAME PIC X(10) VALUE "Old".
   05 WS-AGE  PIC 9(3)  VALUE 99.
PROCEDURE DIVISION.
    INITIALIZE WS-REC.
    DISPLAY WS-NAME.
    DISPLAY WS-AGE.
    STOP RUN.

