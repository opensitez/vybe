*> vybe-test: cobol/cobol/level_88_conditions
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. COND88.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-STATUS PIC X(1) VALUE "A".
   88 IS-ACTIVE  VALUE "A".
   88 IS-INACTIVE VALUE "I".
PROCEDURE DIVISION.
    IF IS-ACTIVE
        DISPLAY "Active"
    END-IF.
    STOP RUN.

