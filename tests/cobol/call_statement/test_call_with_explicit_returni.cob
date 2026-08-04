*> vybe-test: cobol/call_statement/test_call_with_explicit_returning_and_exception
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.
01 WS-RET PIC 9(3) VALUE 0.
01 WS-ARG PIC 9(3) VALUE 55.
PROCEDURE DIVISION.

    CALL "SUBPROG" RETURNING WS-RET
        USING WS-ARG
        ON EXCEPTION
            DISPLAY WS-ARG
            DISPLAY WS-RET
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
    STOP RUN.

