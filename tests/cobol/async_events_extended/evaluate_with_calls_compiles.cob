*> vybe-test: cobol/async_events_extended/evaluate_with_calls_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
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

