*> vybe-test: cobol/module_program_linkage/call_with_on_exception_handling_compiles
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. MAIN-PROG-EX.
PROCEDURE DIVISION.
    CALL "SUBFAIL"
        ON EXCEPTION DISPLAY "FAIL"
        NOT ON EXCEPTION DISPLAY "OK"
    END-CALL.
    STOP RUN.

