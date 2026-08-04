*> vybe-test: cobol/io_and_misc/open_output
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    OPEN OUTPUT WS-FILE.
    STOP RUN.

