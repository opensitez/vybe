*> vybe-test: cobol/debugging_mode/debugging_on_all_references_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG3.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL REFERENCES OF WS-X.
END DECLARATIVES.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-X PIC 9.
PROCEDURE DIVISION.
    STOP RUN.

