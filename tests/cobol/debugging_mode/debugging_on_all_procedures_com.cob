*> vybe-test: cobol/debugging_mode/debugging_on_all_procedures_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG4.
PROCEDURE DIVISION.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL PROCEDURES.
END DECLARATIVES.
    DISPLAY "RUN".
    STOP RUN.

