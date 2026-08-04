*> vybe-test: cobol/io_and_misc/close_and_reopen_file_sequence
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    OPEN OUTPUT WS-FILE.
    CLOSE WS-FILE.
    OPEN INPUT WS-FILE.
    CLOSE WS-FILE.
    STOP RUN.

