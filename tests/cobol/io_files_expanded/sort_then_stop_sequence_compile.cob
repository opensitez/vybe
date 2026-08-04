*> vybe-test: cobol/io_files_expanded/sort_then_stop_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-KEY PIC 9(5).
PROCEDURE DIVISION.
    SORT WS-FILE ON ASCENDING KEY WS-KEY.
    STOP RUN.

