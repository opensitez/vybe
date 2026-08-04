*> vybe-test: cobol/io_files_expanded/write_file_with_from_clause_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-BUF PIC X(20) VALUE "line".
PROCEDURE DIVISION.
    WRITE WS-REC FROM WS-BUF.
    STOP RUN.

