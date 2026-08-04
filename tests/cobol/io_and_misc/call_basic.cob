*> vybe-test: cobol/io_and_misc/call_basic
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9(5).
PROCEDURE DIVISION.
    CALL "SUBPROG" USING X.
    STOP RUN.

