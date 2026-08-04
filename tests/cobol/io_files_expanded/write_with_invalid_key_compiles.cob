*> vybe-test: cobol/io_files_expanded/write_with_invalid_key_compiles
*> origin: languages/cobol/tests/cobol/test_io_files_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    WRITE WS-REC
        INVALID KEY DISPLAY "ERR"
    END-WRITE.
    STOP RUN.

