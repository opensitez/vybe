*> vybe-test: cobol/async_threads/call_with_args
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-INPUT PIC X(20) VALUE "Data".
PROCEDURE DIVISION.
    CALL "SUBB" USING WS-INPUT.
    STOP RUN.

