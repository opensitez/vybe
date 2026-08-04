*> vybe-test: cobol/debugging_mode/debugging_on_procedure_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG2.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON P1.
END DECLARATIVES.
PROCEDURE DIVISION.
P1. STOP RUN.

