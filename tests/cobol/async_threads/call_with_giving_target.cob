*> vybe-test: cobol/async_threads/call_with_giving_target
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-ARG PIC 9(4) VALUE 7.
01 WS-RET PIC 9(4).
PROCEDURE DIVISION.
    CALL "RET-SUB" USING WS-ARG GIVING WS-RET.
    STOP RUN.

