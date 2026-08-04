*> vybe-test: cobol/cobol/if_else
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. IFELSE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-AGE PIC 9(3) VALUE 25.
PROCEDURE DIVISION.
    IF WS-AGE >= 18
        DISPLAY "Adult"
    ELSE
        DISPLAY "Minor"
    END-IF.
    STOP RUN.

