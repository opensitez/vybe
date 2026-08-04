*> vybe-test: cobol/io_files_expanded/merge_file_on_key_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-KEY PIC 9(5).
PROCEDURE DIVISION.
    MERGE WS-FILE ON DESCENDING KEY WS-KEY.
    STOP RUN.

