*> vybe-test: cobol/io_and_misc/accept_from_day
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 D PIC X(5).
PROCEDURE DIVISION.
    ACCEPT D FROM DAY.
    STOP RUN.

