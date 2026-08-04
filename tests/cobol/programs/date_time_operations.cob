*> vybe-test: cobol/programs/date_time_operations
*> origin: languages/cobol/tests/cobol/test_programs.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. DATETIME.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-DATE PIC X(21).
01 WS-TIME PIC X(21).
PROCEDURE DIVISION.
    MOVE FUNCTION CURRENT-DATE TO WS-DATE.
    ACCEPT WS-TIME FROM TIME.
    DISPLAY "Current Date: " WS-DATE.
    DISPLAY "Current Time: " WS-TIME.
    STOP RUN.

