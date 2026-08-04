*> vybe-test: cobol/module_calls_program_linkage/call_exception_runtime_message
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "M7"
        ON EXCEPTION
            DISPLAY "E"
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL
    STOP RUN.

