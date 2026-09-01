*> vybe-test: cobol/async_events_extended/perform_until_with_call_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_async_events_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. C-D.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 2
        ADD 1 TO I
        CALL "SUBD"
    END-PERFORM.
    STOP RUN.

