*> vybe-test: cobol/io_and_misc/call_with_returning_local
*> origin: languages/cobol/tests/cobol/test_io_and_misc.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 RET PIC 9(3).
PROCEDURE DIVISION.
    CALL "SUBPROG" RETURNING RET.
    STOP RUN.

