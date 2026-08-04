*> vybe-test: cobol/module_program_linkage/nested_program_end_program_integration_compiles
*> origin: languages/cobol/tests/cobol/test_module_program_linkage.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. OUTER-MODULE.
PROCEDURE DIVISION.
    DISPLAY "OUTER".
    STOP RUN.
END PROGRAM OUTER-MODULE.

