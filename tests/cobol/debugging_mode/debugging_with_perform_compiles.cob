*> vybe-test: cobol/debugging_mode/debugging_with_perform_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG10.
PROCEDURE DIVISION.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL PROCEDURES.
END DECLARATIVES.
    PERFORM P1.
    STOP RUN.
P1. DISPLAY "P1".

