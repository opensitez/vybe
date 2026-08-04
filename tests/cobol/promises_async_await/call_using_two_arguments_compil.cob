*> vybe-test: cobol/promises_async_await/call_using_two_arguments_compiles
*> origin: languages/cobol/tests/cobol/test_promises_async_await.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 A PIC X(10) VALUE "ALICE".
01 B PIC 9(4) VALUE 1001.
PROCEDURE DIVISION.
    CALL "SUB-C" USING A B.
    STOP RUN.

