*> vybe-test: cobol/module_calls_program_linkage/call_in_perform_loop_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    PERFORM 2 TIMES
        CALL "D"
    END-PERFORM.
    STOP RUN.

