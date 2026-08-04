*> vybe-test: cobol/io_and_misc/accept_basic
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC X(20).
PROCEDURE DIVISION.
    ACCEPT X.
    STOP RUN.

