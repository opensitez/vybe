*> vybe-test: cobol/io_and_misc/sort_descending
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    SORT WS-FILE ON DESCENDING KEY WS-KEY.
    STOP RUN.

