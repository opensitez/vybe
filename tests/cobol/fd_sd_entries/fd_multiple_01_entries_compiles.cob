*> vybe-test: cobol/fd_sd_entries/fd_multiple_01_entries_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
FILE SECTION.
FD F.
01 R1 PIC X(20).
01 R2 PIC X(30).
PROCEDURE DIVISION.
    STOP RUN.

