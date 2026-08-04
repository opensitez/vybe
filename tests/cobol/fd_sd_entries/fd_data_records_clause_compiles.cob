*> vybe-test: cobol/fd_sd_entries/fd_data_records_clause_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
FILE SECTION.
FD F DATA RECORDS ARE R1 R2.
01 R1 PIC X(20).
01 R2 PIC X(30).
PROCEDURE DIVISION.
    STOP RUN.

