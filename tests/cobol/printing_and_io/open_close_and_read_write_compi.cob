*> vybe-test: cobol/printing_and_io/open_close_and_read_write_compiles
*> origin: languages/cobol/tests/cobol/test_printing_and_io.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-REC PIC X(80).
PROCEDURE DIVISION.
    OPEN INPUT WS-FILE.
    READ WS-FILE INTO WS-REC.
    WRITE WS-REC FROM WS-REC.
    CLOSE WS-FILE.
    STOP RUN.

