*> vybe-test: cobol/async_threads/call_with_nested_procedure_text
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-INPUT PIC 9(2) VALUE 10.
PROCEDURE DIVISION.
    IF WS-INPUT > 0
        CALL "SUBX" USING BY VALUE WS-INPUT
    END-IF.
    STOP RUN.

