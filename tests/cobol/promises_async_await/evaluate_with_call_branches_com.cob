*> vybe-test: cobol/promises_async_await/evaluate_with_call_branches_compiles
*> origin: languages/cobol/tests/cobol/test_promises_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE K
        WHEN 1 CALL "P1"
        WHEN 2 CALL "P2"
        WHEN OTHER CALL "PX"
    END-EVALUATE.
    STOP RUN.

