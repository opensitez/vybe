*> vybe-test: cobol/io_files_expanded/rewrite_file_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    REWRITE WS-REC.
    STOP RUN.

