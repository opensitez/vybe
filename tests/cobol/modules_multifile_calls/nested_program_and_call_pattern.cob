*> vybe-test: cobol/modules_multifile_calls/nested_program_and_call_pattern_compiles
*> origin: languages/cobol/tests/cobol/test_modules_multifile_calls.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. OUTER-PROG.
PROCEDURE DIVISION.
    CALL "INNER-PROG".
    STOP RUN.
END PROGRAM OUTER-PROG.

