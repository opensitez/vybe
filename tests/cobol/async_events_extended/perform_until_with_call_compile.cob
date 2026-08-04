*> vybe-test: cobol/async_events_extended/perform_until_with_call_compiles
*> origin: languages/cobol/tests/cobol/test_async_events_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. C-D.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 2
        ADD 1 TO I
        CALL "SUBD"
    END-PERFORM.
    STOP RUN.

