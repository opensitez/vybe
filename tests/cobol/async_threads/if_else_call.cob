*> vybe-test: cobol/async_threads/if_else_call
*> origin: languages/cobol/tests/cobol/test_async_threads.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 F PIC 9 VALUE 1.
PROCEDURE DIVISION.
    IF F = 1
        CALL "YESP"
    ELSE
        CALL "NOP"
    END-IF.
    STOP RUN.

