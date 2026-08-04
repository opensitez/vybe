*> vybe-test: cobol/modules_multifile_calls/call_external_program_no_args_compiles
*> origin: languages/cobol/tests/cobol/test_modules_multifile_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-A.
PROCEDURE DIVISION.
    CALL "SUB-A".
    STOP RUN.

