*> vybe-test: cobol/io_and_misc/sort_ascending
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    SORT WS-FILE ON ASCENDING KEY WS-KEY.
    STOP RUN.

