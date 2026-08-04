*> vybe-test: cobol/network_library_calls/file_open_read_close_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_network_library_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80).
PROCEDURE DIVISION.
    OPEN INPUT WS-FILE.
    READ WS-FILE INTO WS-REC.
    CLOSE WS-FILE.
    STOP RUN.

