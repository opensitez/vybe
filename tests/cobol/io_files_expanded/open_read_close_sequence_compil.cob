*> vybe-test: cobol/io_files_expanded/open_read_close_sequence_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    OPEN INPUT WS-FILE.
    READ WS-FILE INTO WS-REC.
    CLOSE WS-FILE.
    STOP RUN.

