*> vybe-test: cobol/fd_sd_entries/sd_data_records_clause_compiles
*> origin: languages/cobol/tests/cobol/test_fd_sd_entries.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
FILE SECTION.
SD SORT-FILE DATA RECORDS ARE SORT-REC.
01 SORT-REC PIC X(40).
PROCEDURE DIVISION.
    STOP RUN.

