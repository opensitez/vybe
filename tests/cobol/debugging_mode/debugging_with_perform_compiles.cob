*> vybe-test: cobol/debugging_mode/debugging_with_perform_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG10.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL PROCEDURES.
END DECLARATIVES.
PROCEDURE DIVISION.
    PERFORM P1.
    STOP RUN.
P1. DISPLAY "P1".

