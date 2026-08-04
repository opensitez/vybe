*> vybe-test: cobol/io_and_misc/read_file
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 REC PIC X(80).
PROCEDURE DIVISION.
    READ WS-FILE INTO REC.
    STOP RUN.

