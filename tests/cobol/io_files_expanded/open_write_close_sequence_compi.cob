*> vybe-test: cobol/io_files_expanded/open_write_close_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    OPEN OUTPUT WS-FILE.
    WRITE WS-REC.
    CLOSE WS-FILE.
    STOP RUN.

