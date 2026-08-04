*> vybe-test: cobol/async_threads/evaluate_dispatch_calls
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC 9 VALUE 1.
PROCEDURE DIVISION.
    EVALUATE K
        WHEN 1 CALL "S1"
        WHEN 2 CALL "S2"
        WHEN OTHER CALL "SX"
    END-EVALUATE.
    STOP RUN.

