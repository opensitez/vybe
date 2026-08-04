*> vybe-test: cobol/call_statement/test_call_exception_runtime
*> origin: languages/cobol/tests/cobol/test_call_statement.rs
IDENTIFICATION DIVISION.
PROGRAM-ID. T.
DATA DIVISION.
WORKING-STORAGE SECTION.

PROCEDURE DIVISION.

    CALL "NONEXIST"
        ON EXCEPTION
            DISPLAY "ERR"
        NOT ON EXCEPTION
            DISPLAY "OK"
    END-CALL.
    STOP RUN.

