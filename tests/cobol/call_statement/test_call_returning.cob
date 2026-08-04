*> vybe-test: cobol/call_statement/test_call_returning
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RET PIC 9(3) VALUE 0.
PROCEDURE DIVISION.

    CALL "SUBPROG" RETURNING WS-RET.
    STOP RUN.

