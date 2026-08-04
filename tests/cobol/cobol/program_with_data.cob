*> vybe-test: cobol/cobol/program_with_data
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. VARS.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-NAME PIC X(20) VALUE "Alice".
01 WS-AGE  PIC 9(3)  VALUE 30.
PROCEDURE DIVISION.
    DISPLAY WS-NAME.
    DISPLAY WS-AGE.
    STOP RUN.

