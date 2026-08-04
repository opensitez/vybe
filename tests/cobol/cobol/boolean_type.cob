*> vybe-test: cobol/cobol/boolean_type
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. BOOL.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FLAG PIC 9(1) VALUE 1.
PROCEDURE DIVISION.
    IF WS-FLAG = 1
        DISPLAY "True"
    ELSE
        DISPLAY "False"
    END-IF.
    STOP RUN.

