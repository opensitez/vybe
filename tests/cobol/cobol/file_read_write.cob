*> vybe-test: cobol/cobol/file_read_write
*> origin: languages/cobol/tests/cobol/test_cobol.rs

IDENTIFICATION DIVISION.
PROGRAM-ID. FILEIO.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RECORD PIC X(80).
PROCEDURE DIVISION.
    DISPLAY "File I/O test".
    STOP RUN.

