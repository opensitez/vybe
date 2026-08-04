*> vybe-test: cobol/fd_sd_entries/fd_block_contains_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
FILE SECTION.
FD F BLOCK CONTAINS 5 RECORDS.
01 R PIC X(80).
PROCEDURE DIVISION.
    STOP RUN.

