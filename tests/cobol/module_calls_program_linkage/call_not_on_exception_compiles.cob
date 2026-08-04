*> vybe-test: cobol/module_calls_program_linkage/call_not_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "M8" NOT ON EXCEPTION DISPLAY "OK" END-CALL.
    STOP RUN.

