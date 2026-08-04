*> vybe-test: cobol/debugging_mode/debugging_with_data_division_compiles
*> origin: languages/cobol/tests/cobol/test_debugging_mode.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. DBG8.
DECLARATIVES.
DBG-SEC SECTION.
    USE FOR DEBUGGING ON ALL PROCEDURES.
END DECLARATIVES.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 X PIC 9.
PROCEDURE DIVISION.
    MOVE 1 TO X.
    STOP RUN.

