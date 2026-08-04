*> vybe-test: cobol/async_threads/perform_times_with_call
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    PERFORM 2 TIMES
        CALL "SUBD"
    END-PERFORM.
    STOP RUN.

