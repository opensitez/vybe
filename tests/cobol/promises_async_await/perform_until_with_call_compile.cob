*> vybe-test: cobol/promises_async_await/perform_until_with_call_compiles
*> origin: languages/cobol/tests/cobol/test_promises_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 I PIC 9 VALUE 0.
PROCEDURE DIVISION.
    PERFORM UNTIL I >= 2
        ADD 1 TO I
        CALL "SUB-L"
    END-PERFORM.
    STOP RUN.

