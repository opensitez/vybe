*> vybe-test: cobol/exceptions_error_paths/call_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "MAYBE"
        ON EXCEPTION DISPLAY "FAIL"
        NOT ON EXCEPTION DISPLAY "OK"
    END-CALL.
    STOP RUN.

