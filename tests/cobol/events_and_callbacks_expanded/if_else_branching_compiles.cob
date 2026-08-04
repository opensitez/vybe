*> vybe-test: cobol/events_and_callbacks_expanded/if_else_branching_compiles
*> origin: languages/cobol/tests/cobol/test_events_and_callbacks_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF F = 1 DISPLAY "Y" ELSE DISPLAY "N" END-IF.
    STOP RUN.

