*> vybe-test: cobol/fd_sd_entries/fd_label_records_standard_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
FILE SECTION.
FD F LABEL RECORDS ARE STANDARD.
01 R PIC X(20).
PROCEDURE DIVISION.
    STOP RUN.

