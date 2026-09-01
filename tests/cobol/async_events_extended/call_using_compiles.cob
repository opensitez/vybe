*> vybe-test: cobol/async_events_extended/call_using_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_async_events_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. C-B.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 V PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    CALL "SUBB" USING V.
    STOP RUN.

