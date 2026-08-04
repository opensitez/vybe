*> vybe-test: cobol/cobol/set_statement
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SETST.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9(1) VALUE 0.
   88 IS-ON  VALUE 1.
   88 IS-OFF VALUE 0.
PROCEDURE DIVISION.
    SET IS-ON TO TRUE.
    DISPLAY WS-FLAG.
    STOP RUN.

