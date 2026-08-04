*> vybe-test: cobol/modules_multifile_calls/call_external_program_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_modules_multifile_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-E.
PROCEDURE DIVISION.
    CALL "SUB-E"
        ON EXCEPTION DISPLAY "FAIL"
        NOT ON EXCEPTION DISPLAY "OK"
    END-CALL.
    STOP RUN.

