*> vybe-test: cobol/enterprise/global_external_program
*> origin: languages/cobol/tests/cobol/test_enterprise.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. SHARED.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-CONFIG PIC X(100) GLOBAL VALUE "production".
01 WS-DB-CONN PIC X(100) EXTERNAL.
01 WS-COUNTER PIC 9(10) GLOBAL VALUE 0.
PROCEDURE DIVISION.
    DISPLAY "Config: " WS-CONFIG.
    ADD 1 TO WS-COUNTER.
    DISPLAY "Counter: " WS-COUNTER.
    STOP RUN.

