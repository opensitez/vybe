*> vybe-test: cobol/debugging_mode/debugging_on_references_second_var_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG7.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-Y PIC 9.
PROCEDURE DIVISION.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL REFERENCES OF WS-Y.
END DECLARATIVES.
    STOP RUN.

