*> vybe-test: cobol/debugging_mode/debugging_nested_procedure_sections_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG11.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON SEC1.
    USE FOR DEBUGGING ON SEC2.
END DECLARATIVES.
PROCEDURE DIVISION.
SEC1 SECTION.
    PERFORM P1.
SEC2 SECTION.
P1. STOP RUN.

