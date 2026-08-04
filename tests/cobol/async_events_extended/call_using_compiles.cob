*> vybe-test: cobol/async_events_extended/call_using_compiles
*> origin: languages/cobol/tests/cobol/test_async_events_extended.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. C-B.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 V PIC 9(3) VALUE 1.
PROCEDURE DIVISION.
    CALL "SUBB" USING V.
    STOP RUN.

