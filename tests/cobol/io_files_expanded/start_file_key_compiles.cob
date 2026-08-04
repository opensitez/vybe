*> vybe-test: cobol/io_files_expanded/start_file_key_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    START WS-FILE KEY IS = WS-REC.
    STOP RUN.

