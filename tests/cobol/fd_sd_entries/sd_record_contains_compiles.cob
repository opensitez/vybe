*> vybe-test: cobol/fd_sd_entries/sd_record_contains_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
FILE SECTION.
SD SORT-FILE RECORD CONTAINS 50 CHARACTERS.
01 SORT-REC PIC X(50).
PROCEDURE DIVISION.
    STOP RUN.

