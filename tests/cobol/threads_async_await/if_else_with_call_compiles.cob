*> vybe-test: cobol/threads_async_await/if_else_with_call_compiles
*> origin: languages/cobol/tests/cobol/test_threads_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF F = 1
        CALL "SUBY"
    ELSE
        CALL "SUBZ"
    END-IF.
    STOP RUN.

