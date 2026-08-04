*> vybe-test: cobol/cobol/func_current_date
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. FDATE.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO WS-DATE.
    DISPLAY WS-DATE.
    STOP RUN.

