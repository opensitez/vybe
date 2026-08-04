*> vybe-test: cobol/io_and_misc/open_input
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    OPEN INPUT WS-FILE.
    STOP RUN.

