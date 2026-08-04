*> vybe-test: cobol/module_calls_program_linkage/call_on_exception_compiles
*> origin: languages/cobol/tests/cobol/test_module_calls_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
PROCEDURE DIVISION.
    CALL "M7" ON EXCEPTION DISPLAY "E" END-CALL.
    STOP RUN.

