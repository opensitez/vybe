*> vybe-test: cobol/exceptions_error_paths/retry_loop_pattern_compiles
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

