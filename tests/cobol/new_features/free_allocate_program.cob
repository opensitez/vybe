*> vybe-test: cobol/new_features/free_allocate_program
*> origin: languages/cobol/tests/cobol/test_new_features.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. DYNALLOC.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-PTR PIC X(100).
PROCEDURE DIVISION.
    ALLOCATE WS-PTR.
    DISPLAY "Allocated".
    FREE WS-PTR.
    DISPLAY "Freed".
    STOP RUN.

