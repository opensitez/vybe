*> vybe-test: cobol/io_files_expanded/open_extend_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    OPEN EXTEND WS-FILE.
    STOP RUN.

