*> vybe-test: cobol/io_and_misc/accept_from_date
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(8).
PROCEDURE DIVISION.
    ACCEPT D FROM DATE.
    STOP RUN.

