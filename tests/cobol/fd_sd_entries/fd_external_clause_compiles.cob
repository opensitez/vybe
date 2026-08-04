*> vybe-test: cobol/fd_sd_entries/fd_external_clause_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
FILE SECTION.
FD F EXTERNAL.
01 R PIC X(20).
PROCEDURE DIVISION.
    STOP RUN.

