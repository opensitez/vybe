*> vybe-test: cobol/async_events_extended/evaluate_with_calls_compiles
*> origin: languages/cobol/tests/cobol/test_async_events_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. C-E.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 K PIC 9 VALUE 2.
PROCEDURE DIVISION.
    EVALUATE K
        WHEN 1 CALL "S1"
        WHEN 2 CALL "S2"
        WHEN OTHER CALL "SX"
    END-EVALUATE.
    STOP RUN.

