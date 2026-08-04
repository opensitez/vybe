*> vybe-test: cobol/debugging_mode/debugging_on_single_section_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG5.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON SEC1.
END DECLARATIVES.
PROCEDURE DIVISION.
SEC1 SECTION.
P1. STOP RUN.

