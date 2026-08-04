*> vybe-test: cobol/events_and_callbacks_expanded/cics_delay_statement_compiles
*> origin: languages/cobol/tests/cobol/test_events_and_callbacks_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    EXEC CICS DELAY SECONDS(1) END-EXEC.
    STOP RUN.

