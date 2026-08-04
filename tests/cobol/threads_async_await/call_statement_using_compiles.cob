*> vybe-test: cobol/threads_async_await/call_statement_using_compiles
*> origin: languages/cobol/tests/cobol/test_threads_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 V PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    CALL "SUBB" USING V.
    STOP RUN.

