*> vybe-test: cobol/open_close_modes/close_with_no_rewind_compiles
*> origin: languages/cobol/tests/cobol/test_open_close_modes.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO "f.dat".
DATA DIVISION.
FILE SECTION.
FD F.
01 R PIC X(20).
PROCEDURE DIVISION.
    OPEN INPUT F.
    CLOSE F WITH NO REWIND.
    STOP RUN.

