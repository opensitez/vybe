*> vybe-test: cobol/threads_async_await/evaluate_with_call_branches_compiles
*> origin: languages/cobol/tests/cobol/test_threads_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC 9 VALUE 1.
PROCEDURE DIVISION.
    EVALUATE K
        WHEN 1 CALL "SUB1"
        WHEN 2 CALL "SUB2"
        WHEN OTHER CALL "SUBX"
    END-EVALUATE.
    STOP RUN.

