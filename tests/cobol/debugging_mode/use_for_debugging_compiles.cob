*> vybe-test: cobol/debugging_mode/use_for_debugging_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG1.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL PROCEDURES.
END DECLARATIVES.
PROCEDURE DIVISION.
    STOP RUN.

