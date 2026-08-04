*> vybe-test: cobol/events_and_callbacks_expanded/declaratives_use_after_error_compiles
*> origin: languages/cobol/tests/cobol/test_events_and_callbacks_expanded.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.
    DECLARATIVES.
    D-SEC SECTION.
        USE AFTER STANDARD ERROR PROCEDURE ON WS-FILE.
    END DECLARATIVES.
    DISPLAY "RUN".
    STOP RUN.

