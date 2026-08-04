*> vybe-test: cobol/async_threads/call_on_exception
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "SUBC"
        ON EXCEPTION DISPLAY "E"
        NOT ON EXCEPTION DISPLAY "O"
    END-CALL.
    STOP RUN.

