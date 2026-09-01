*> vybe-test: cobol/exceptions_error_paths/retry_loop_pattern_compiles
*> vybe-test-mode: compile
*> `CALL "…"` names a program that does not exist in this run unit. cobc
*> compiles it and then aborts — `libcob: error: module not found` — so
*> "runs and exits 0" is not a property this source has under any COBOL.
*> What it CAN assert is the one its name claims: that it compiles.
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 N PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL N >= 3
        ADD 1 TO N
        CALL "TRY-STEP"
    END-PERFORM.
    STOP RUN.

