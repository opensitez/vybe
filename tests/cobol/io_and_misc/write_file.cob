*> vybe-test: cobol/io_and_misc/write_file
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC PIC X(80) VALUE "Data".
PROCEDURE DIVISION.
    WRITE WS-REC FROM REC.
    STOP RUN.

