*> vybe-test: cobol/io_and_misc/sort_and_merge_program
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-KEY PIC 9(5).
PROCEDURE DIVISION.
    SORT WS-FILE ON ASCENDING KEY WS-KEY.
    MERGE WS-FILE ON DESCENDING KEY WS-KEY.
    STOP RUN.

