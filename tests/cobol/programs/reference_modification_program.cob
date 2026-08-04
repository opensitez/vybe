*> vybe-test: cobol/programs/reference_modification_program
*> origin: languages/cobol/tests/cobol/test_programs.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. REFMOD.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-FULL-NAME PIC X(30) VALUE "John Michael Smith".
01 WS-FIRST     PIC X(10).
01 WS-MIDDLE    PIC X(10).
01 WS-LAST      PIC X(10).
PROCEDURE DIVISION.
    MOVE WS-FULL-NAME(1:4) TO WS-FIRST.
    MOVE WS-FULL-NAME(6:7) TO WS-MIDDLE.
    MOVE WS-FULL-NAME(14:5) TO WS-LAST.
    DISPLAY "First:  " WS-FIRST.
    DISPLAY "Middle: " WS-MIDDLE.
    DISPLAY "Last:   " WS-LAST.
    STOP RUN.

