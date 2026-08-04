*> vybe-test: cobol/exceptions_error_paths/cancel_on_error_compiles
*> origin: languages/cobol/tests/cobol/test_exceptions_error_paths.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    CALL "MAYBE"
        ON EXCEPTION CANCEL "MAYBE"
    END-CALL.
    STOP RUN.

