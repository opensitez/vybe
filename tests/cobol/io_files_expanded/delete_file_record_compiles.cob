*> vybe-test: cobol/io_files_expanded/delete_file_record_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    DELETE WS-FILE.
    STOP RUN.

