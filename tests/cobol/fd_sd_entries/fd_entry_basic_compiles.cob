*> vybe-test: cobol/fd_sd_entries/fd_entry_basic_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
ENVIRONMENT DIVISION.
INPUT-OUTPUT SECTION.
FILE-CONTROL.
    SELECT F ASSIGN TO "f.dat".
DATA DIVISION.
FILE SECTION.
FD F RECORD CONTAINS 80 CHARACTERS.
01 R PIC X(80).
PROCEDURE DIVISION.
    STOP RUN.

