*> vybe-test: cobol/fd_sd_entries/sd_entry_basic_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
FILE SECTION.
SD SORT-FILE.
01 SORT-REC.
   05 SORT-KEY PIC 9(5).
PROCEDURE DIVISION.
    STOP RUN.

