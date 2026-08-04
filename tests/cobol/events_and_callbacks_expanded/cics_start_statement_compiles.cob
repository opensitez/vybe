*> vybe-test: cobol/events_and_callbacks_expanded/cics_start_statement_compiles
*> origin: languages/cobol/tests/cobol/test_events_and_callbacks_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC CICS START TRANSID(NXTT) END-EXEC.
    STOP RUN.

