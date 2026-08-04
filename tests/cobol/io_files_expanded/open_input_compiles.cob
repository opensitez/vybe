*> vybe-test: cobol/io_files_expanded/open_input_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    OPEN INPUT WS-FILE.
    STOP RUN.

