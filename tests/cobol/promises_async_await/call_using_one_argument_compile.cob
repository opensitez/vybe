*> vybe-test: cobol/promises_async_await/call_using_one_argument_compiles
*> origin: languages/cobol/tests/cobol/test_promises_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9(3) VALUE 7.
PROCEDURE DIVISION.
    CALL "SUB-B" USING N.
    STOP RUN.

