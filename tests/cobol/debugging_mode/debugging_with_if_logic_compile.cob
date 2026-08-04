*> vybe-test: cobol/debugging_mode/debugging_with_if_logic_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG9.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL PROCEDURES.
END DECLARATIVES.
PROCEDURE DIVISION.
    IF 1 = 1 DISPLAY "Y" END-IF.
    STOP RUN.

