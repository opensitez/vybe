*> vybe-test: cobol/io_and_misc/accept_from_time
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 T PIC X(8).
PROCEDURE DIVISION.
    ACCEPT T FROM TIME.
    STOP RUN.

