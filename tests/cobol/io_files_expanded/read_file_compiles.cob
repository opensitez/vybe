*> vybe-test: cobol/io_files_expanded/read_file_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    READ WS-FILE INTO WS-REC.
    STOP RUN.

