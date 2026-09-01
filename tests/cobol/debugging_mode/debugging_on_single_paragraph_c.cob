*> vybe-test: cobol/debugging_mode/debugging_on_single_paragraph_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG6.
PROCEDURE DIVISION.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON P1.
END DECLARATIVES.
P1. DISPLAY "P".
    STOP RUN.

